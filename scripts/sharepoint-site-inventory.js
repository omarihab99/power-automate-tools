/* eslint-disable */
// SharePoint site inventory - READ ONLY.
//
// Paste this whole file into the browser console on any page of the target
// SharePoint site (it reuses your signed-in session cookies, no token needed),
// then run:
//
//   await SPInventory.run()                    everything, returns the result object
//   SPInventory.download()                     save the last result as JSON
//   SPInventory.downloadCsv()                  save a flat "who has what where" CSV
//
//   await SPInventory.steps.permissionLevels() one step on its own
//   await SPInventory.steps.groups()
//   await SPInventory.steps.lists()
//
//   SPInventory.config.siteUrl = 'https://tenant.sharepoint.com/sites/x'   another site / subsite
//   SPInventory.config.includeHiddenLists = true     include system/hidden lists
//   SPInventory.config.includeHiddenFields = true    include hidden columns
//   SPInventory.config.includeLimitedAccess = true   keep "Limited Access" assignments
//   SPInventory.config.includeGroupMembers = false   skip reading group members
//   SPInventory.config.scanItemPermissions = 'all'   also scan list items ('libraries' default, 'none' to skip)
//   SPInventory.config.maxItemsToScan = 50000        skip the scan for bigger lists/libraries
//
// Lists and document libraries are both covered. Each one gets its columns,
// content types and permissions; libraries also get versioning / check-out /
// approval settings, size and file + folder counts, and the folders and files
// that have their own (broken-inheritance) permissions.
//
// Only GET requests are sent: nothing on the site is created, changed or deleted.
// Reading role assignments needs the "Enumerate Permissions" right (site owner
// or admin); without it the permission parts are reported as errors and the
// rest of the inventory still completes.

window.SPInventory = window.SPInventory || {};

(function (SPI) {
  'use strict';

  SPI.config = Object.assign(
    {
      siteUrl: null,
      includeHiddenLists: false,
      includeHiddenFields: false,
      includeLimitedAccess: false,
      includeGroupMembers: true,
      // Find folders/files/items with broken inheritance: 'libraries', 'all' or 'none'.
      scanItemPermissions: 'libraries',
      maxItemsToScan: 50000,
      concurrency: 4,
    },
    SPI.config || {}
  );

  // ---------------------------------------------------------------- logging

  const log = {
    step: (m) => console.log('%c== ' + m, 'color:#0078d4;font-weight:bold'),
    info: (m) => console.log('   ' + m),
    ok: (m) => console.log('%c   OK ' + m, 'color:#107c10'),
    warn: (m) => console.warn('   ' + m),
    fail: (m) => console.error('   FAIL ' + m),
  };
  SPI.log = log;

  // ---------------------------------------------------------------- http

  function webUrl() {
    if (SPI.config.siteUrl) {
      return SPI.config.siteUrl.replace(/\/+$/, '');
    }
    const ctx = window._spPageContextInfo;
    if (ctx && ctx.webAbsoluteUrl) {
      return ctx.webAbsoluteUrl.replace(/\/+$/, '');
    }
    // Fallback: site collection root (/sites/x or /teams/x). Subsites need config.siteUrl.
    const m = /^\/(?:sites|teams)\/[^/]+/i.exec(location.pathname);
    return location.origin + (m ? m[0] : '');
  }
  SPI.webUrl = webUrl;

  async function get(path) {
    const url = /^https?:\/\//i.test(path) ? path : webUrl() + '/_api/' + path.replace(/^\/+/, '');
    const response = await fetch(url, {
      method: 'GET',
      credentials: 'include',
      headers: { Accept: 'application/json;odata=nometadata' },
    });
    const text = await response.text();
    let body = null;
    try {
      body = text ? JSON.parse(text) : null;
    } catch (_) {
      body = text;
    }
    if (!response.ok) {
      const message =
        (body && body['odata.error'] && body['odata.error'].message && body['odata.error'].message.value) ||
        (body && body.error && body.error.message) ||
        (typeof body === 'string' ? body.slice(0, 300) : '') ||
        response.statusText;
      const error = new Error(response.status + ': ' + message);
      error.status = response.status;
      throw error;
    }
    return body;
  }

  // Follows odata.nextLink until every page is read.
  async function getAll(path) {
    const items = [];
    let next = path;
    while (next) {
      const page = await get(next);
      items.push(...((page && page.value) || []));
      next = (page && (page['odata.nextLink'] || page['@odata.nextLink'])) || null;
    }
    return items;
  }

  SPI.get = get;
  SPI.getAll = getAll;

  async function mapLimit(items, limit, fn) {
    const results = new Array(items.length);
    let index = 0;
    const worker = async () => {
      while (index < items.length) {
        const i = index++;
        results[i] = await fn(items[i], i);
      }
    };
    await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
    return results;
  }

  // ---------------------------------------------------------------- permissions

  // SP.PermissionKind: value n is bit (n - 1) of the 64-bit Low/High mask.
  const PERMISSION_KINDS = {
    1: 'ViewListItems', 2: 'AddListItems', 3: 'EditListItems', 4: 'DeleteListItems',
    5: 'ApproveItems', 6: 'OpenItems', 7: 'ViewVersions', 8: 'DeleteVersions',
    9: 'CancelCheckout', 10: 'ManagePersonalViews', 12: 'ManageLists', 13: 'ViewFormPages',
    14: 'AnonymousSearchAccessList', 17: 'Open', 18: 'ViewPages', 19: 'AddAndCustomizePages',
    20: 'ApplyThemeAndBorder', 21: 'ApplyStyleSheets', 22: 'ViewUsageData', 23: 'CreateSSCSite',
    24: 'ManageSubwebs', 25: 'CreateGroups', 26: 'ManagePermissions', 27: 'BrowseDirectories',
    28: 'BrowseUserInfo', 29: 'AddDelPrivateWebParts', 30: 'UpdatePersonalWebParts',
    31: 'ManageWeb', 32: 'AnonymousSearchAccessWebLists', 37: 'UseClientIntegration',
    38: 'UseRemoteAPIs', 39: 'ManageAlerts', 40: 'CreateAlerts', 41: 'EditMyUserInfo',
    63: 'EnumeratePermissions',
  };

  function decodeBasePermissions(bp) {
    if (!bp) {
      return [];
    }
    const low = Number(bp.Low) >>> 0;
    const high = Number(bp.High) >>> 0;
    if (low === 0xffffffff && (high & 0x7fffffff) === 0x7fffffff) {
      return ['FullMask'];
    }
    const names = [];
    for (const [kind, name] of Object.entries(PERMISSION_KINDS)) {
      const bit = Number(kind) - 1;
      const set = bit < 32 ? (low >>> bit) & 1 : (high >>> (bit - 32)) & 1;
      if (set) {
        names.push(name);
      }
    }
    return names;
  }
  SPI.decodeBasePermissions = decodeBasePermissions;

  const PRINCIPAL_TYPES = { 1: 'User', 2: 'DistributionList', 4: 'SecurityGroup', 8: 'SharePointGroup' };

  function principal(member) {
    return {
      id: member.Id,
      title: member.Title,
      loginName: member.LoginName,
      type: PRINCIPAL_TYPES[member.PrincipalType] || String(member.PrincipalType),
      email: member.Email || undefined,
    };
  }

  // RoleTypeKind 1 = Guest ("Limited Access"), granted automatically by item-level sharing.
  function isLimitedAccess(roleDef) {
    return roleDef.RoleTypeKind === 1 || roleDef.Name === 'Limited Access';
  }

  async function getRoleAssignments(basePath) {
    const raw = await getAll(basePath + '/roleassignments?$expand=Member,RoleDefinitionBindings');
    return raw
      .map((ra) => {
        const levels = (ra.RoleDefinitionBindings || []).filter(
          (rd) => SPI.config.includeLimitedAccess || !isLimitedAccess(rd)
        );
        return {
          principal: principal(ra.Member),
          permissionLevels: levels.map((rd) => rd.Name),
        };
      })
      .filter((ra) => ra.permissionLevels.length);
  }

  // ---------------------------------------------------------------- steps

  SPI.steps = {};
  SPI.result = SPI.result || {};

  SPI.steps.web = async function () {
    log.step('Site');
    const web = await get(
      'web?$select=Id,Title,Url,ServerRelativeUrl,Description,Created,WebTemplate,Language,HasUniqueRoleAssignments'
    );
    const result = {
      id: web.Id,
      title: web.Title,
      url: web.Url,
      description: web.Description,
      created: web.Created,
      template: web.WebTemplate,
      language: web.Language,
      hasUniquePermissions: web.HasUniqueRoleAssignments,
    };
    try {
      result.permissions = await getRoleAssignments('web');
      log.ok(result.permissions.length + ' role assignment(s) on the site');
    } catch (e) {
      result.permissionsError = e.message;
      log.warn('site permissions: ' + e.message);
    }
    SPI.result.site = result;
    return result;
  };

  SPI.steps.permissionLevels = async function () {
    log.step('Permission levels');
    const defs = await getAll('web/roledefinitions');
    const result = defs
      .filter((d) => SPI.config.includeLimitedAccess || !isLimitedAccess(d))
      .sort((a, b) => a.Order - b.Order)
      .map((d) => ({
        id: d.Id,
        name: d.Name,
        description: d.Description,
        hidden: d.Hidden,
        roleTypeKind: d.RoleTypeKind,
        custom: d.RoleTypeKind === 0,
        permissions: decodeBasePermissions(d.BasePermissions),
        basePermissions: d.BasePermissions,
      }));
    console.table(result.map((d) => ({ name: d.name, custom: d.custom, permissions: d.permissions.length, description: d.description })));
    SPI.result.permissionLevels = result;
    return result;
  };

  SPI.steps.groups = async function () {
    log.step('SharePoint groups');
    const groups = await getAll(
      'web/sitegroups?$select=Id,Title,Description,LoginName,OwnerTitle,AllowMembersEditMembership,' +
        'AllowRequestToJoinLeave,OnlyAllowMembersViewMembership,RequestToJoinLeaveEmailSetting'
    );

    let webAssignments = SPI.result.site && SPI.result.site.permissions;
    if (!webAssignments) {
      try {
        webAssignments = await getRoleAssignments('web');
      } catch (e) {
        log.warn('site permissions: ' + e.message);
        webAssignments = [];
      }
    }

    const result = await mapLimit(groups, SPI.config.concurrency, async (g) => {
      const group = {
        id: g.Id,
        title: g.Title,
        description: g.Description,
        owner: g.OwnerTitle,
        onlyMembersCanViewMembership: g.OnlyAllowMembersViewMembership,
        membersCanEditMembership: g.AllowMembersEditMembership,
        allowRequestToJoinLeave: g.AllowRequestToJoinLeave,
        requestEmail: g.RequestToJoinLeaveEmailSetting || undefined,
        sitePermissionLevels: webAssignments
          .filter((ra) => ra.principal.id === g.Id)
          .flatMap((ra) => ra.permissionLevels),
        listPermissions: [],
      };
      if (SPI.config.includeGroupMembers) {
        try {
          const users = await getAll('web/sitegroups(' + g.Id + ')/users?$select=Id,Title,LoginName,Email,PrincipalType');
          group.members = users.map(principal);
        } catch (e) {
          group.membersError = e.message;
        }
      }
      return group;
    });

    console.table(
      result.map((g) => ({
        group: g.title,
        sitePermissionLevels: g.sitePermissionLevels.join(', ') || '(none on site)',
        members: g.members ? g.members.length : g.membersError || '-',
        owner: g.owner,
      }))
    );
    SPI.result.groups = result;
    linkGroupsToLists();
    return result;
  };

  function mapField(f) {
    const field = {
      id: f.Id,
      title: f.Title,
      internalName: f.InternalName,
      type: f.TypeAsString,
      required: f.Required,
      readOnly: f.ReadOnlyField,
      hidden: f.Hidden,
      indexed: f.Indexed,
      unique: f.EnforceUniqueValues,
      fromBaseType: f.FromBaseType,
      group: f.Group,
      description: f.Description || undefined,
      defaultValue: f.DefaultValue || undefined,
    };
    if (f.Choices) field.choices = f.Choices;
    if (f.FillInChoice !== undefined && f.Choices) field.allowFillIn = f.FillInChoice;
    if (f.LookupList) {
      field.lookupList = f.LookupList;
      field.lookupField = f.LookupField;
    }
    if (f.AllowMultipleValues !== undefined) field.allowMultipleValues = f.AllowMultipleValues;
    if (f.SelectionMode !== undefined) field.selectionMode = f.SelectionMode === 0 ? 'PeopleOnly' : 'PeopleAndGroups';
    if (f.MaxLength !== undefined) field.maxLength = f.MaxLength;
    if (f.NumberOfLines !== undefined) field.numberOfLines = f.NumberOfLines;
    if (f.RichText !== undefined) field.richText = f.RichText;
    if (f.MinimumValue !== undefined) {
      field.min = f.MinimumValue;
      field.max = f.MaximumValue;
    }
    if (f.DisplayFormat !== undefined) field.displayFormat = f.DisplayFormat;
    if (f.Formula) field.formula = f.Formula;
    if (f.CurrencyLocaleId !== undefined) field.currencyLocaleId = f.CurrencyLocaleId;
    if (f.ValidationFormula) {
      field.validationFormula = f.ValidationFormula;
      field.validationMessage = f.ValidationMessage;
    }
    return field;
  }

  const DRAFT_VISIBILITY = { 0: 'AnyReader', 1: 'Author', 2: 'Approver' };

  function shouldScanItems(list) {
    const mode = SPI.config.scanItemPermissions;
    return mode === 'all' || (mode === 'libraries' && list.type === 'DocumentLibrary');
  }

  // Walks every item/file/folder (5000 per page) to count files and folders and
  // to find the ones with broken inheritance, then reads their role assignments.
  async function scanItems(base, list) {
    if (list.itemCount > SPI.config.maxItemsToScan) {
      list.itemScanSkipped =
        'itemCount ' + list.itemCount + ' > config.maxItemsToScan ' + SPI.config.maxItemsToScan;
      return;
    }
    const items = await getAll(
      base + '/items?$select=Id,FileRef,FileLeafRef,FSObjType,HasUniqueRoleAssignments&$top=5000'
    );
    if (list.type === 'DocumentLibrary') {
      list.folderCount = items.filter((i) => i.FSObjType === 1).length;
      list.fileCount = items.length - list.folderCount;
    }
    const unique = items.filter((i) => i.HasUniqueRoleAssignments);
    list.uniquePermissionItems = await mapLimit(unique, SPI.config.concurrency, async (i) => {
      const entry = {
        id: i.Id,
        type: i.FSObjType === 1 ? 'Folder' : list.type === 'DocumentLibrary' ? 'File' : 'Item',
        name: i.FileLeafRef,
        path: i.FileRef,
      };
      try {
        entry.permissions = await getRoleAssignments(base + '/items(' + i.Id + ')');
      } catch (e) {
        entry.permissionsError = e.message;
      }
      return entry;
    });
  }

  SPI.steps.lists = async function () {
    log.step('Lists and document libraries');
    const filter = SPI.config.includeHiddenLists ? '' : '&$filter=Hidden eq false';
    const lists = await getAll(
      'web/lists?$select=Id,Title,Description,BaseTemplate,BaseType,Hidden,ItemCount,Created,' +
        'LastItemModifiedDate,HasUniqueRoleAssignments,EnableVersioning,EnableAttachments,' +
        'EnableMinorVersions,MajorVersionLimit,MajorWithMinorVersionsLimit,DraftVersionVisibility,' +
        'EnableModeration,ForceCheckout,DocumentTemplateUrl,ContentTypesEnabled,EnableFolderCreation,' +
        'DefaultViewUrl,RootFolder/ServerRelativeUrl&$expand=RootFolder' + filter
    );
    const libraryCount = lists.filter((l) => l.BaseType === 1).length;
    log.info(lists.length - libraryCount + ' list(s), ' + libraryCount + ' document library(ies)');

    let done = 0;
    const result = await mapLimit(lists, SPI.config.concurrency, async (l) => {
      const base = "web/lists(guid'" + l.Id + "')";
      const isLibrary = l.BaseType === 1;
      const list = {
        id: l.Id,
        title: l.Title,
        description: l.Description || undefined,
        url: l.RootFolder && l.RootFolder.ServerRelativeUrl,
        defaultViewUrl: l.DefaultViewUrl,
        type: isLibrary ? 'DocumentLibrary' : 'List',
        baseTemplate: l.BaseTemplate,
        hidden: l.Hidden,
        itemCount: l.ItemCount,
        created: l.Created,
        lastItemModified: l.LastItemModifiedDate,
        versioning: l.EnableVersioning,
        majorVersionLimit: l.EnableVersioning ? l.MajorVersionLimit : undefined,
        contentApproval: l.EnableModeration,
        folderCreation: l.EnableFolderCreation,
        contentTypesEnabled: l.ContentTypesEnabled,
        hasUniquePermissions: l.HasUniqueRoleAssignments,
      };
      if (isLibrary) {
        Object.assign(list, {
          minorVersions: l.EnableMinorVersions,
          minorVersionLimit: l.EnableMinorVersions ? l.MajorWithMinorVersionsLimit : undefined,
          draftVisibility: l.EnableMinorVersions || l.EnableModeration ? DRAFT_VISIBILITY[l.DraftVersionVisibility] : undefined,
          requireCheckout: l.ForceCheckout,
          documentTemplateUrl: l.DocumentTemplateUrl || undefined,
        });
        try {
          const folder = await get(base + '/rootfolder?$select=StorageMetrics&$expand=StorageMetrics');
          const m = folder && folder.StorageMetrics;
          if (m) {
            list.totalSizeBytes = Number(m.TotalSize);
            list.totalSizeMB = Math.round((Number(m.TotalSize) / 1048576) * 100) / 100;
            list.totalFileCount = Number(m.TotalFileCount);
          }
        } catch (e) {
          list.storageError = e.message;
        }
      } else {
        list.attachments = l.EnableAttachments;
      }

      try {
        const fieldFilter = SPI.config.includeHiddenFields ? '' : '?$filter=Hidden eq false';
        list.columns = (await getAll(base + '/fields' + fieldFilter)).map(mapField);
      } catch (e) {
        list.columnsError = e.message;
      }

      try {
        const cts = await getAll(base + '/contenttypes?$select=Id,Name,Description,Group,Hidden');
        list.contentTypes = cts
          .filter((ct) => SPI.config.includeHiddenFields || !ct.Hidden)
          .map((ct) => ({
            id: (ct.Id && ct.Id.StringValue) || ct.StringId,
            name: ct.Name,
            group: ct.Group,
            description: ct.Description || undefined,
          }));
      } catch (e) {
        list.contentTypesError = e.message;
      }

      if (l.HasUniqueRoleAssignments) {
        try {
          list.permissions = await getRoleAssignments(base);
        } catch (e) {
          list.permissionsError = e.message;
        }
      } else {
        list.permissionsInheritedFrom = webUrl();
      }

      if (shouldScanItems(list)) {
        try {
          await scanItems(base, list);
        } catch (e) {
          list.itemScanError = e.message;
        }
      }

      const hasError =
        list.columnsError || list.permissionsError || list.contentTypesError || list.storageError || list.itemScanError;
      log.info('[' + ++done + '/' + lists.length + '] ' + l.Title + (hasError ? ' (with errors)' : ''));
      return list;
    });

    const permissionSummary = (l) =>
      l.hasUniquePermissions
        ? l.permissions
          ? l.permissions.map((p) => p.principal.title + ': ' + p.permissionLevels.join('/')).join('; ')
          : l.permissionsError
        : 'Inherited';
    const customColumns = (l) =>
      l.columns ? l.columns.filter((c) => !c.fromBaseType || c.internalName === 'Title').length : l.columnsError;
    const uniqueItems = (l) =>
      l.uniquePermissionItems ? l.uniquePermissionItems.length : l.itemScanSkipped || l.itemScanError ? 'not scanned' : '-';

    const plainLists = result.filter((l) => l.type === 'List');
    const libraries = result.filter((l) => l.type === 'DocumentLibrary');
    if (plainLists.length) {
      log.info('Lists:');
      console.table(
        plainLists.map((l) => ({
          list: l.title,
          items: l.itemCount,
          columns: customColumns(l),
          versioning: l.versioning,
          permissions: permissionSummary(l),
          uniqueItemPermissions: uniqueItems(l),
        }))
      );
    }
    if (libraries.length) {
      log.info('Document libraries:');
      console.table(
        libraries.map((l) => ({
          library: l.title,
          url: l.url,
          files: l.fileCount != null ? l.fileCount : l.totalFileCount,
          folders: l.folderCount,
          sizeMB: l.totalSizeMB,
          columns: customColumns(l),
          versioning: l.versioning ? 'major' + (l.minorVersions ? '+minor' : '') + ' (' + l.majorVersionLimit + ')' : 'off',
          checkout: l.requireCheckout,
          approval: l.contentApproval,
          permissions: permissionSummary(l),
          uniqueFileFolderPermissions: uniqueItems(l),
        }))
      );
    }
    SPI.result.lists = result;
    linkGroupsToLists();
    return result;
  };

  // Adds to each group the lists/libraries, and the folders/files/items inside
  // them, where it has unique (non-inherited) permissions.
  function linkGroupsToLists() {
    const { groups, lists } = SPI.result;
    if (!groups || !lists) {
      return;
    }
    for (const group of groups) {
      group.listPermissions = [];
      group.itemPermissions = [];
      for (const list of lists) {
        for (const ra of list.permissions || []) {
          if (ra.principal.id === group.id) {
            group.listPermissions.push({ list: list.title, type: list.type, permissionLevels: ra.permissionLevels });
          }
        }
        for (const item of list.uniquePermissionItems || []) {
          for (const ra of item.permissions || []) {
            if (ra.principal.id === group.id) {
              group.itemPermissions.push({ list: list.title, type: item.type, path: item.path, permissionLevels: ra.permissionLevels });
            }
          }
        }
      }
    }
  }

  // ---------------------------------------------------------------- entry points

  const ORDER = ['web', 'permissionLevels', 'groups', 'lists'];

  SPI.run = async function () {
    const started = Date.now();
    SPI.result = { generatedAt: new Date().toISOString(), siteUrl: webUrl() };
    log.step('SharePoint inventory (read only) -> ' + webUrl());

    for (const name of ORDER) {
      try {
        await SPI.steps[name]();
      } catch (error) {
        SPI.result[name + 'Error'] = error.message;
        log.fail('step "' + name + '" failed: ' + error.message);
        console.error(error);
      }
    }

    log.step('Done in ' + Math.round((Date.now() - started) / 1000) + 's - SPInventory.result holds the data');
    log.info('SPInventory.download() for JSON, SPInventory.downloadCsv() for a flat permissions CSV');
    return SPI.result;
  };

  function save(content, fileName, type) {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(new Blob([content], { type }));
    link.download = fileName;
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
  }

  function siteSlug() {
    return (webUrl().split('/').pop() || 'site').replace(/[^\w.-]+/g, '_');
  }

  SPI.download = function (fileName) {
    save(JSON.stringify(SPI.result, null, 2), fileName || 'sp-inventory-' + siteSlug() + '.json', 'application/json');
  };

  // One row per (scope, principal, permission level).
  SPI.permissionRows = function () {
    const rows = [];
    const add = (scope, scopeName, ra) =>
      ra.permissionLevels.forEach((level) =>
        rows.push({
          scope,
          scopeName,
          principal: ra.principal.title,
          principalType: ra.principal.type,
          loginName: ra.principal.loginName,
          permissionLevel: level,
        })
      );
    const { site, lists } = SPI.result;
    ((site && site.permissions) || []).forEach((ra) => add('Site', site.title, ra));
    (lists || []).forEach((l) => {
      const scope = l.type === 'DocumentLibrary' ? 'Library' : 'List';
      if (l.hasUniquePermissions) {
        (l.permissions || []).forEach((ra) => add(scope, l.title, ra));
      } else {
        rows.push({ scope, scopeName: l.title, principal: '(inherits from site)', principalType: '', loginName: '', permissionLevel: '' });
      }
      (l.uniquePermissionItems || []).forEach((item) =>
        (item.permissions || []).forEach((ra) => add(item.type, item.path, ra))
      );
    });
    return rows;
  };

  SPI.downloadCsv = function (fileName) {
    const rows = SPI.permissionRows();
    const headers = ['scope', 'scopeName', 'principal', 'principalType', 'loginName', 'permissionLevel'];
    const escape = (v) => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
    const csv = [headers.join(',')].concat(rows.map((r) => headers.map((h) => escape(r[h])).join(','))).join('\r\n');
    save('\uFEFF' + csv, fileName || 'sp-permissions-' + siteSlug() + '.csv', 'text/csv');
  };

  console.log('SPInventory loaded for ' + webUrl() + ' - run: await SPInventory.run()');
})(window.SPInventory);
