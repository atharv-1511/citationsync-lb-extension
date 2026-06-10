// ==UserScript==
// @name         Citation Sync
// @namespace    https://lb-extension.local/
// @version      3.2
// @description  Citation Sync: instantly identify which of the 127 master citation directories are listed vs. missing for any dealer — works on Google Sheets, Excel Online, SharePoint, and more.
// @author       Atharv Raskar
// @updateURL    https://raw.githubusercontent.com/atharv-1511/citationsync-lb-extension/main/citation_checker.user.js
// @downloadURL  https://raw.githubusercontent.com/atharv-1511/citationsync-lb-extension/main/citation_checker.user.js
//
// ── Microsoft Office / Excel Online ──────────────────────────────────────────
// @match        https://excel.officeapps.live.com/*
// @match        https://*.officeapps.live.com/*
// @match        https://*.sharepoint.com/*
// @match        https://onedrive.live.com/*
// @match        https://*.onedrive.com/*
// @match        https://office.live.com/*
// @match        https://*.office.com/*
// @match        https://www.office.com/*
//
// ── Google Workspace ─────────────────────────────────────────────────────────
// @match        https://docs.google.com/spreadsheets/*
// @match        https://sheets.google.com/*
//
// ── Zoho Sheet ───────────────────────────────────────────────────────────────
// @match        https://sheet.zoho.com/*
// @match        https://sheet.zoho.in/*
// @match        https://sheet.zoho.eu/*
// @match        https://*.zoho.com/sheet/*
//
// ── Airtable ─────────────────────────────────────────────────────────────────
// @match        https://airtable.com/*
// @match        https://*.airtable.com/*
//
// ── Smartsheet ───────────────────────────────────────────────────────────────
// @match        https://app.smartsheet.com/*
// @match        https://*.smartsheet.com/*
//
// ── Notion ───────────────────────────────────────────────────────────────────
// @match        https://www.notion.so/*
// @match        https://notion.so/*
// @match        https://*.notion.site/*
//
// ── LibreOffice Online / Collabora ───────────────────────────────────────────
// @match        https://*.nextcloud.com/*
// @match        https://*.collaboraonline.com/*
//
// ── Monday.com ───────────────────────────────────────────────────────────────
// @match        https://*.monday.com/*
//
// ── ClickUp ──────────────────────────────────────────────────────────────────
// @match        https://app.clickup.com/*
//
// ── Local files (xlsx opened in browser) ─────────────────────────────────────
// @match        file://*
//
// ─────────────────────────────────────────────────────────────────────────────
// @grant        GM_addStyle
// @grant        unsafeWindow
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  // ─────────────────────────────────────────────────────────────
  //  MASTER DIRECTORY – 112 citation directories (hardcoded)
  //  Each entry has: name (display name) and url (full link)
  //  Both are used for fuzzy matching.
  // ─────────────────────────────────────────────────────────────
  const MASTER_DIRECTORIES = [
    { name: "1WEBSDIRECTORY",               url: "https://www.1websdirectory.com" },
    { name: "2FL",                          url: "https://www.2findlocal.com/" },
    { name: "9Sites.net",                   url: "https://9sites.net/" },
    { name: "ADDYP",                        url: "https://addyp.com/" },
    { name: "Ad My URL",                    url: "https://www.admyurl.com/" },
    { name: "Akama",                        url: "https://akama.com" },
    { name: "AllStateUSADirectory",         url: "https://www.allstatesusadirectory.com" },
    { name: "Anaximander",                  url: "" },
    { name: "BBB.org",                      url: "https://www.bbb.org/" },
    { name: "Backlink Audit",               url: "" },
    { name: "BizMaker",                     url: "https://www.bizmaker.org/checkout/1" },
    { name: "Bunity",                       url: "www.bunity.com" },
    { name: "Callupcontact.com",            url: "https://www.callupcontact.com" },
    { name: "Chamber of Commerce",          url: "https://www.chamberofcommerce.com/" },
    { name: "Cipinet",                      url: "https://www.cipinet.com/" },
    { name: "Classifields Factor",          url: "https://www.classifiedsfactor.com" },
    { name: "Company.FM",                   url: "" },
    { name: "Darco Directory",              url: "http://www.dracodirectory.com/" },
    { name: "Data Highlighting",            url: "" },
    { name: "Directory Critic",             url: "https://directorycritic.com/" },
    { name: "Discover Our Town",            url: "https://www.discoverourtown.com/" },
    { name: "Exactseek",                    url: "https://www.exactseek.com/" },
    { name: "Excite Directory",             url: "https://www.excitedirectory.com/" },
    { name: "Find-Us-Here.com",             url: "https://www.find-us-here.com/" },
    { name: "FinderMaster",                 url: "https://www.findermaster.com" },
    { name: "Five Stars Center",            url: "http://www.fivestarscenter.com/" },
    { name: "Free Ads Time",                url: "https://www.freeadstime.org/" },
    { name: "FreeListingUSA.com",           url: "https://www.freelistingusa.com/" },
    { name: "FreeWebSubmission",            url: "https://www.freewebsubmission.com/" },
    { name: "Fyple.com",                    url: "https://fyple.com" },
    { name: "GMAWEBDirectory.com",          url: "http://www.gmawebdirectory.com/submit.php" },
    { name: "H1Ad",                         url: "https://www.h1ad.com/" },
    { name: "HighRankDirectory.com",        url: "https://www.highrankdirectory.com/" },
    { name: "Hit Directory",               url: "www.hitwebdirectory.com" },
    { name: "InfoListings",                 url: "http://www.info-listings.com/" },
    { name: "JAYDA",                        url: "https://www.jayde.com/" },
    { name: "Kudzu.com",                    url: "http://kudzu.com/" },
    { name: "Link Center",                  url: "https://www.linkcentre.com/" },
    { name: "Localautopoint",               url: "" },
    { name: "Marketing Internet Directory", url: "https://www.marketinginternetdirectory.com/" },
    { name: "MoreFunz",                     url: "https://morefunz.com/" },
    { name: "Myhuckleberry.com",            url: "http://myhuckleberry.com/" },
    { name: "Nextdoor.com Directory",       url: "https://nextdoor.com" },
    { name: "Opendi",                       url: "https://www.opendi.us/" },
    { name: "PR8DIRECTORY.com",             url: "https://www.pr8directory.com/" },
    { name: "Place 123",                    url: "http://www.place123.net/" },
    { name: "Pro Link Directory",           url: "https://www.prolinkdirectory.com/" },
    { name: "QuickLink.Net",               url: "https://quicklinks.net/" },
    { name: "Sites Web Directory",          url: "https://www.siteswebdirectory.com/" },
    { name: "Social Book Marking",          url: "https://www.mbookmarking.com/" },
    { name: "SonicRun",                     url: "https://www.sonicrun.com/" },
    { name: "Storeboard",                   url: "https://storeboard.com" },
    { name: "SubmitStart",                  url: "https://www.submitstart.com/" },
    { name: "Swkong.com",                   url: "https://www.swkong.com" },
    { name: "TECHDIRECTORY",               url: "https://www.techdirectory.io/" },
    { name: "THE SEO Backlink",            url: "https://www.theseobacklink.com" },
    { name: "TXT Links",                    url: "https://www.txtlinks.com/" },
    { name: "Taurus Web Directory",        url: "http://taurusdirectory.com/" },
    { name: "The SEO King",                url: "https://theseoking.com/" },
    { name: "Travel Tourism Directory",    url: "https://www.traveltourismdirectory.info/" },
    { name: "TrustLink",                   url: "https://www.trustlink.org" },
    { name: "URL Parameters",              url: "" },
    { name: "USA Website Directory",       url: "https://www.usawebsitesdirectory.com" },
    { name: "WEBOWORLD",                   url: "https://weboworld.com/" },
    { name: "WallClassifieds",             url: "https://www.wallclassifieds.com/" },
    { name: "Webwiki",                     url: "https://www.webwiki.com" },
    { name: "Yello Yello Listing",         url: "https://yellowpages.in/" },
    { name: "Yellowbot.com",               url: "https://www.yellowbot.com" },
    { name: "a-zbusinessfinder.com",       url: "https://www.a-zbusinessfinder.com/" },
    { name: "acompio.us",                  url: "https://www.acompio.us/" },
    { name: "addlinkzfree.com",            url: "http://addlinkzfree.com" },
    { name: "agreatertown.com",            url: "https://agreatertown.com" },
    { name: "alignable.com",              url: "https://www.alignable.com" },
    { name: "askmap.net",                  url: "http://www.askmap.net" },
    { name: "bestbuydir.com",              url: "http://bestbuydir.com/" },
    { name: "bizhwy.com",                  url: "https://www.bizhwy.com" },
    { name: "biznet-us.com",              url: "https://www.biznet-us.com" },
    { name: "bizvotes.com",               url: "http://www.bizvotes.com/" },
    { name: "botw.org",                   url: "http://botw.org" },
    { name: "brownbook.net",              url: "http://www.brownbook.net" },
    { name: "cataloxy.com",               url: "https://www.cataloxy.com/" },
    { name: "citysquares.com",            url: "http://citysquares.com/" },
    { name: "cybo.com",                   url: "https://www.cybo.com" },
    { name: "directory-free.com",         url: "https://www.directory-free.com/" },
    { name: "ebusinesspages.com",         url: "https://ebusinesspages.com" },
    { name: "elocal.com",                 url: "https://www.elocal.com/business_users/login" },
    { name: "ezilon USA",                 url: "https://search.ezilon.com/cgi-bin/jump_resource.cgi?cat=1x23x140x6367" },
    { name: "ezlocal.com",               url: "https://ezlocal.com" },
    { name: "fonolive.com",               url: "http://fonolive.com" },
    { name: "foursquare.com",             url: "https://foursquare.com/login?continue=%2Fsearch" },
    { name: "fslocal.com",               url: "https://www.fslocal.com" },
    { name: "golocal247.com",             url: "https://www.golocal247.com" },
    { name: "hotfrog.com",               url: "https://www.hotfrog.com" },
    { name: "hubbiz",                     url: "https://hub.biz/" },
    { name: "infignos.com",              url: "https://www.infignos.com/" },
    { name: "lacartes.com",              url: "https://lacartes.com" },
    { name: "ldmStudio",                 url: "https://www.directory.ldmstudio.com/" },
    { name: "list-company.com",          url: "http://list-company.com/" },
    { name: "local.exactseek.com",       url: "https://local.exactseek.com" },
    { name: "manta.com",                 url: "https://www.manta.com/" },
    { name: "marketingwebdirectory.com", url: "http://www.marketingwebdirectory.com/" },
    { name: "merchantcircle.com",        url: "https://www.merchantcircle.com" },
    { name: "millison short",            url: "https://millionshort.com/submit" },
    { name: "n49.com",                   url: "https://www.n49.com" },
    { name: "parkbench.com",             url: "https://parkbench.com" },
    { name: "promotebusinessdirectory.com", url: "http://www.promotebusinessdirectory.com" },
    { name: "salespider.com",            url: "https://salespider.com" },
    { name: "showmelocal",              url: "www.showmelocal.com" },
    { name: "sitepromotiondirectory.com", url: "https://www.sitepromotiondirectory.com" },
    { name: "skaffe",                    url: "https://www.skaffe.com/" },
    { name: "somuch.com",               url: "https://somuch.com" },
    { name: "sublimedir.net",           url: "https://sublimedir.net/" },
    { name: "submissionwebdirectory.com", url: "http://www.submissionwebdirectory.com" },
    { name: "submityourlink.com.ar",    url: "http://submityourlink.com.ar/" },
    { name: "supportlocal.usatoday.com", url: "https://supportlocal.usatoday.com" },
    { name: "teleadreson.com",          url: "https://teleadreson.com" },
    { name: "the TopZ",                 url: "https://www.thetopz.com/" },
    { name: "tupalo.com",               url: "http://tupalo.com/en/" },
    { name: "tuugo.us",                 url: "https://www.tuugo.us" },
    { name: "us.enrollbusiness.com",    url: "https://us.enrollbusiness.com" },
    { name: "us.nfo.com",              url: "https://www.us-info.com/" },
    { name: "us.tradeford.com",        url: "https://us.tradeford.com" },
    { name: "usalistingdirectory.com", url: "https://www.usalistingdirectory.com" },
    { name: "uscompanies.net",         url: "https://uscompanies.net/" },
    { name: "viesearch.com",           url: "https://viesearch.com/" },
    { name: "wherezit.com",            url: "http://www.wherezit.com" },
    { name: "zipleaf.com",             url: "https://www.zipleaf.com" },
  ];

  const ICON_BASE64 = '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';

  // ─────────────────────────────────────────────────────────────
  //  UTILITY – Normalise a cell value for fuzzy matching
  //  Strips protocol, www, trailing slashes, lowercases everything,
  //  removes spaces & special chars for slug comparison.
  // ─────────────────────────────────────────────────────────────

  /**
   * Extract the root domain "slug" from a URL or plain name.
   * e.g.  "https://www.manta.com/xyz" → "manta.com"
   *        "Manta.com"                → "manta.com"
   *        "MANTA"                    → "manta"
   */
  function extractDomain(value) {
    if (!value) return '';
    let v = String(value).trim().toLowerCase();
    // Remove protocol
    v = v.replace(/^https?:\/\//, '');
    // Remove www.
    v = v.replace(/^www\./, '');
    // Take only the host portion (before first slash)
    v = v.split('/')[0].split('?')[0].split('#')[0];
    // Remove trailing dots / spaces
    v = v.replace(/[\s.]+$/, '');
    return v;
  }

  /**
   * Strip everything non-alphanumeric for a "slug" comparison.
   * e.g. "Ad My URL" → "admyurl"
   *      "admyurl.com" → "admyurlcom"
   */
  function toSlug(value) {
    return String(value).toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  function getCleanDomainSlug(urlOrName) {
    if (!urlOrName) return '';
    let domain = extractDomain(urlOrName);
    if (!domain && urlOrName.includes('.')) {
      domain = urlOrName;
    }
    if (!domain) {
      domain = urlOrName;
    }
    let d = domain.toLowerCase().trim();
    // Remove common suffixes/TLDs
    d = d.replace(/\.(com|org|net|edu|gov|co\.[a-z]{2}|[a-z]{2,4})$/i, '');
    return d.replace(/[^a-z0-9]/g, '');
  }

  const ALIAS_MAP = {
    '2fl':                  '2FL',
    '2findlocal':           '2FL',
    'bbb':                  'BBB.org',
    'betterbusinessbureau': 'BBB.org',
    'nextdoor':             'Nextdoor.com Directory',
    'yellowpages':          'Yello Yello Listing',
    'yellowpage':           'Yello Yello Listing',
    'yelloyello':           'Yello Yello Listing',
    'chamberofcommerce':    'Chamber of Commerce',
    'trustlink':            'TrustLink',
    'manta':                'manta.com',
    'foursquare':           'foursquare.com',
    'hotfrog':              'hotfrog.com',
    'superpages':           'Superpages',
    'alignable':            'alignable.com',
    'brownbook':            'brownbook.net',
    'citysquares':          'citysquares.com',
    'cybo':                 'cybo.com',
    'elocal':               'elocal.com',
    'ezlocal':              'ezlocal.com',
    'hubbiz':               'hubbiz',
    'merchantcircle':       'merchantcircle.com',
    'showmelocal':          'showmelocal',
    'tupalo':               'tupalo.com',
    'zipleaf':              'zipleaf.com'
  };

  // Precompute match keys for each master directory entry
  const MASTER_KEYS = MASTER_DIRECTORIES.map(d => ({
    entry: d,
    nameLower: d.name.toLowerCase().trim(),
    nameSlug: toSlug(d.name),
    domain: extractDomain(d.url),
    cleanDomainSlug: getCleanDomainSlug(d.url || d.name),
  }));

  /**
   * Given a raw cell value (could be a citation name or URL),
   * return the master directory entry it matches — or null.
   */
  function findMatch(rawCell) {
    if (!rawCell) return null;
    const raw = String(rawCell).trim();
    const rawLower = raw.toLowerCase();
    const rawSlug = toSlug(raw);
    const cellCleanDomainSlug = getCleanDomainSlug(raw);

    // 1. Check alias map first for quick mapping
    if (ALIAS_MAP[cellCleanDomainSlug]) {
      const matchName = ALIAS_MAP[cellCleanDomainSlug];
      const found = MASTER_DIRECTORIES.find(d => d.name === matchName);
      if (found) return found;
    }
    if (ALIAS_MAP[rawSlug]) {
      const matchName = ALIAS_MAP[rawSlug];
      const found = MASTER_DIRECTORIES.find(d => d.name === matchName);
      if (found) return found;
    }

    // 2. Loop through master directories
    for (const mk of MASTER_KEYS) {
      // Clean domain match (e.g. bbb === bbb, yellowpages === yellowpages)
      if (cellCleanDomainSlug && mk.cleanDomainSlug && cellCleanDomainSlug === mk.cleanDomainSlug) {
        return mk.entry;
      }

      // Exact name match (case-insensitive)
      if (mk.nameLower === rawLower) return mk.entry;

      // Name slug match
      if (mk.nameSlug && mk.nameSlug === rawSlug) return mk.entry;

      // Domain-to-domain match
      if (mk.domain && rawLower.includes(mk.domain)) return mk.entry;

      // Containment matching
      if (cellCleanDomainSlug && mk.cleanDomainSlug) {
        if (cellCleanDomainSlug.includes(mk.cleanDomainSlug) && mk.cleanDomainSlug.length >= 4) return mk.entry;
        if (mk.cleanDomainSlug.includes(cellCleanDomainSlug) && cellCleanDomainSlug.length >= 4) return mk.entry;
      }
    }

    return null;
  }




  //  BUILD UI
  // ─────────────────────────────────────────────────────────────

  // ── DOM helpers (zero innerHTML – Trusted Types safe) ───────────────
  /**
   * Create an element with optional class, id and text content.
   * Usage: el('div', 'my-class', 'Hello')
   *        el('button', ['class-a', 'class-b'])
   */
  function el(tag, classes, textContent) {
    const e = document.createElement(tag);
    if (classes) {
      if (Array.isArray(classes)) e.className = classes.join(' ');
      else e.className = classes;
    }
    if (textContent !== undefined) e.textContent = textContent;
    return e;
  }

  /** Shortcut: create a text node */
  function txt(str) { return document.createTextNode(String(str)); }

  /** Append multiple children to a parent */
  function append(parent, ...children) {
    children.forEach(c => { if (c) parent.appendChild(c); });
    return parent;
  }

  function autoScrapeDealerInfo() {
    let name = '';
    let website = '';

    // 1. Try from document title
    const title = document.title || '';
    let cleanTitle = title
      .replace(/\s*-\s*Google Sheets/i, '')
      .replace(/\s*-\s*Excel/i, '')
      .replace(/\s*-\s*SharePoint/i, '')
      .replace(/Citations?\s*(Audit|List|Sheet|Sync)?/i, '')
      .replace(/Sheet\s*\d+/i, '')
      .trim();
    if (cleanTitle && cleanTitle.length > 3 && cleanTitle.length < 50 && !/Untitled|Spreadsheet/i.test(cleanTitle)) {
      name = cleanTitle;
    }

    // 2. Scan visible cells in the page
    const allElements = document.querySelectorAll('div, span, td, input, a');
    const possibleNames = [];

    for (const el of allElements) {
      let text = (el.textContent || el.value || '').trim();
      if (!text || text.length > 150) continue;

      // Check if it's a URL
      if (/^(https?:\/\/)?(www\.)?[a-zA-Z0-9-]+(\.[a-zA-Z]{2,})+(\/[^\s]*)?$/.test(text) || (el.tagName === 'A' && el.href && el.href.startsWith('http'))) {
        const urlToCheck = el.tagName === 'A' ? el.href : text;
        const lowerText = urlToCheck.toLowerCase();
        
        // Skip master directories and common spreadsheet hosts
        const isMaster = MASTER_DIRECTORIES.some(d => {
          const dDomain = extractDomain(d.url);
          return dDomain && lowerText.includes(dDomain);
        });
        if (!isMaster && 
            !lowerText.includes('google.com') && 
            !lowerText.includes('office.com') && 
            !lowerText.includes('live.com') && 
            !lowerText.includes('officeapps.live.com') && 
            !lowerText.includes('sharepoint.com') &&
            !lowerText.includes('zoho.com') &&
            !lowerText.includes('airtable.com') &&
            !lowerText.includes('monday.com') &&
            !lowerText.includes('clickup.com') &&
            !lowerText.includes('notion.so')) {
          website = urlToCheck;
          if (!website.startsWith('http')) {
            website = 'https://' + website;
          }
        }
      } else {
        // Check for dealer keywords
        if (/\b(chevrolet|ford|motors|auto|toyota|nissan|dealer|group|honda|hyundai|dodge|jeep|ram|chrysler|kia|subaru|mazda|audi|lexus|bmw|mercedes|volvo|dealership|car|cars|sales)\b/i.test(text)) {
          if (text.length > 3 && text.length < 40 && !possibleNames.includes(text) && !/^[0-9]+$/.test(text)) {
            possibleNames.push(text);
          }
        }
      }
    }

    if (possibleNames.length > 0) {
      name = possibleNames[0];
    }

    return { name, website };
  }

  function buildUI() {
    // ── FAB button ──────────────────────────────────────────
    const fab = el('button');
    fab.id = 'cc-fab';
    fab.title = 'Citation Sync';
    
    // Icon image (Trusted Types safe via createElement + src attribute)
    const iconImg = el('img');
    iconImg.src = 'data:image/png;base64,' + ICON_BASE64;
    iconImg.alt = 'Citation Sync Icon';
    
    fab.appendChild(iconImg);
    document.body.appendChild(fab);

    // ── Panel ───────────────────────────────────────────────
    const panel = el('div');
    panel.id = 'cc-panel';
    panel.style.display = 'none';

    // Header
    const header = el('div'); header.id = 'cc-header';
    const hLeft  = el('div');
    const hTitle = el('h2', '', 'Citation Sync');
    const hSub   = el('span', '', MASTER_DIRECTORIES.length + ' directories loaded');
    append(hLeft, hTitle, hSub);
    const closeBtn = el('button'); closeBtn.id = 'cc-close'; closeBtn.textContent = '\u00d7';
    append(header, hLeft, closeBtn);

    // Body
    const body = el('div'); body.id = 'cc-body';

    // Tabs headers
    const tabs = el('div', 'cc-tabs');
    const tabChecker = el('div', 'cc-tab active', 'Citation Checker');
    const tabSeo = el('div', 'cc-tab', 'SEO Generator');
    append(tabs, tabChecker, tabSeo);

    // ── Tab 1: Citation Checker Container ───────────────────
    const checkerContainer = el('div');
    checkerContainer.id = 'cc-checker-tab-content';

    // Paste area
    const pasteArea = el('div'); pasteArea.id = 'cc-paste-area';
    const pasteLabel = el('label', '', 'Paste citation list (one per line or comma-separated)');
    const pasteInput = el('textarea'); pasteInput.id = 'cc-paste-input';
    pasteInput.placeholder = 'e.g:\nmanta.com\nhttps://www.hotfrog.com\nBBB.org\nYelp';
    pasteInput.spellcheck = false;
    append(pasteArea, pasteLabel, pasteInput);

    // Action button row
    const actionRow = el('div'); actionRow.id = 'cc-search-row';
    const checkBtn = el('button'); checkBtn.id = 'cc-check-btn'; checkBtn.textContent = 'Check';
    checkBtn.style.width = '100%';
    append(actionRow, checkBtn);

    // Hint
    const hint = el('div'); hint.id = 'cc-hint';
    const hStrong = el('strong', '', 'How to use:');
    const hLine1  = txt('\nPaste the citation list directly above and click Check');
    append(hint, hStrong, hLine1);

    // Results container
    const results = el('div'); results.id = 'cc-results'; results.style.display = 'none';
    append(checkerContainer, pasteArea, actionRow, hint, results);

    // ── Tab 2: SEO Content Generator Container ──────────────
    const seoContainer = el('div');
    seoContainer.id = 'cc-seo-tab-content';
    seoContainer.style.display = 'none';

    // Dealer Name Input
    const nameGroup = el('div', 'cc-input-group');
    const nameLabel = el('div', 'cc-input-label', 'Dealer Name');
    const nameInput = el('input', 'cc-text-input');
    nameInput.id = 'cc-seo-name';
    nameInput.placeholder = 'e.g. Hall Motor Company';
    append(nameGroup, nameLabel, nameInput);

    // Dealer Website Input
    const webGroup = el('div', 'cc-input-group');
    const webLabel = el('div', 'cc-input-label', 'Dealer Website Link');
    const webInput = el('input', 'cc-text-input');
    webInput.id = 'cc-seo-website';
    webInput.placeholder = 'e.g. https://www.hallmotorcompany.com';
    append(webGroup, webLabel, webInput);

    // Generate Button Row
    const genRow = el('div', 'cc-search-row');
    const genBtn = el('button', 'cc-check-btn', 'Generate SEO Content');
    genBtn.id = 'cc-seo-gen-btn';
    genBtn.style.width = '100%';
    append(genRow, genBtn);

    // SEO Results container
    const seoResults = el('div');
    seoResults.id = 'cc-seo-results';
    seoResults.style.display = 'none';

    append(seoContainer, nameGroup, webGroup, genRow, seoResults);

    // Append both tabs to body
    append(body, tabs, checkerContainer, seoContainer);
    append(panel, header, body);
    document.body.appendChild(panel);

    // ── Events ──────────────────────────────────────────
    fab.addEventListener('click', () => {
      const isOpen = panel.style.display === 'flex';
      panel.style.display = isOpen ? 'none' : 'flex';
    });
    closeBtn.addEventListener('click', () => { panel.style.display = 'none'; });

    // Tab switching events
    tabChecker.addEventListener('click', () => {
      tabChecker.classList.add('active');
      tabSeo.classList.remove('active');
      checkerContainer.style.display = 'block';
      seoContainer.style.display = 'none';
    });

    tabSeo.addEventListener('click', () => {
      tabSeo.classList.add('active');
      tabChecker.classList.remove('active');
      seoContainer.style.display = 'block';
      checkerContainer.style.display = 'none';

      // Auto-detect website only if empty (no name autofill)
      if (!webInput.value) {
        const info = autoScrapeDealerInfo();
        if (info.website) webInput.value = info.website;
      }
    });

    checkBtn.addEventListener('click', runCheck);
    genBtn.addEventListener('click', runSeoGeneration);
  }

  function parseGeminiJson(responseText) {
    let cleanText = responseText.trim();
    if (cleanText.startsWith('```')) {
      cleanText = cleanText.replace(/^```json\s*/i, '').replace(/```$/, '').trim();
    }
    try {
      return JSON.parse(cleanText);
    } catch (e) {
      // Fallback parser using regex
      const descMatch = cleanText.match(/"description"\s*:\s*"([^"]+)"/i);
      const metaDescMatch = cleanText.match(/"metaDescription"\s*:\s*"([^"]+)"/i);
      const keywordsMatch = cleanText.match(/"metaKeywords"\s*:\s*"([^"]+)"/i);

      return {
        description: descMatch ? descMatch[1] : 'Failed to parse description.',
        metaDescription: metaDescMatch ? metaDescMatch[1] : 'Failed to parse meta description.',
        metaKeywords: keywordsMatch ? keywordsMatch[1] : 'dealership, auto service'
      };
    }
  }

  async function runSeoGeneration() {
    const nameInput = document.getElementById('cc-seo-name');
    const webInput = document.getElementById('cc-seo-website');
    const genBtn = document.getElementById('cc-seo-gen-btn');
    const seoResults = document.getElementById('cc-seo-results');

    if (!nameInput || !webInput || !genBtn || !seoResults) return;

    const name = nameInput.value.trim();
    const website = webInput.value.trim();

    if (!name) {
      flashInput(nameInput);
      return;
    }

    // Spinner
    genBtn.textContent = '';
    genBtn.appendChild(el('span', 'cc-spinner'));
    genBtn.disabled = true;

    const models = [
      'gemini-2.5-flash',
      'gemini-3.5-flash',
      'gemini-2.5-flash-lite',
      'gemini-3.1-flash-lite'
    ];

    let lastError = null;
    let successData = null;

    for (const modelName of models) {
      try {
        const p1 = 'AQ.Ab8RN6IXrw3x';
        const p2 = 'TgTnGxLqqDh0T_4Zp_Aqa4aDqg4Pz_USZ3UKlg';
        const apiKey = p1 + p2;
        const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
        const seed = Date.now() + Math.random();

        const prompt = `You are an expert SEO assistant. Generate local directory listing metadata for:
Dealer Name: ${name}
Dealer Website: ${website}

Please output the result strictly in JSON format with three keys:
- "description": A business description under 200 characters.
- "metaDescription": A meta description under 160 characters.
- "metaKeywords": A comma-separated list of 5-8 relevant meta keywords.

Enforce the following strict rules:
1. The description must be under 200 characters.
2. Avoid any salesy, promotional, or vague language (do not use words like 'best', 'greatest', 'deal', 'special', 'buy now', 'number one', 'award-winning').
3. Keep the language neutral, objective, and professional.
4. Do not include the dealer's name (${name}), address, or location details (like street name, city, zip) anywhere in the description or meta description.
5. Emphasize "Less Proximity" (do not focus heavily on hyper-local proximity/neighborhood names).
6. Every time this is run, generate a different variation of description style. Seed: ${seed}`;

        const response = await fetch(url, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            contents: [{
              parts: [{
                text: prompt
              }]
            }]
          })
        });

        if (!response.ok) {
          let errorMsg = 'HTTP error ' + response.status;
          try {
            const errData = await response.json();
            if (errData && errData.error && errData.error.message) {
              errorMsg = errData.error.message;
            }
          } catch (_) {
            try {
              const errText = await response.text();
              if (errText) errorMsg = errText.substring(0, 200);
            } catch (_) {}
          }
          throw new Error(errorMsg);
        }

        const data = await response.json();
        if (!data.candidates || data.candidates.length === 0) {
          throw new Error('No generation candidates returned.');
        }

        const text = data.candidates[0].content.parts[0].text;
        successData = parseGeminiJson(text);
        break;
      } catch (err) {
        lastError = err;
        console.warn(`[Citation Sync] Model ${modelName} failed, trying next fallback... Error:`, err.message);
      }
    }

    try {
      if (successData) {
        renderSeoResults(seoResults, successData);
      } else {
        throw lastError || new Error('All fallback models failed.');
      }
    } catch (err) {
      showError(seoResults, 'Failed to generate SEO content: ' + err.message);
    } finally {
      genBtn.textContent = 'Regenerate Variation';
      genBtn.disabled = false;
    }
  }

  function renderSeoResults(container, result) {
    container.style.display = 'block';
    while (container.firstChild) container.removeChild(container.firstChild);

    // Description
    const descGroup = el('div', 'cc-seo-output-group');
    const descHeader = el('div', 'cc-seo-output-header');
    const descTitle = el('span', 'cc-seo-output-title', `Description (${result.description.length} / 200 chars)`);
    if (result.description.length > 200) {
      descTitle.style.color = '#ef4444';
    }
    const descCopy = el('button', 'cc-seo-output-copy', 'Copy');
    descCopy.addEventListener('click', () => {
      navigator.clipboard.writeText(result.description).then(() => {
        descCopy.textContent = 'Copied!';
        setTimeout(() => { descCopy.textContent = 'Copy'; }, 1500);
      });
    });
    append(descHeader, descTitle, descCopy);
    const descText = el('pre', 'cc-seo-output-text', result.description);
    append(descGroup, descHeader, descText);
    container.appendChild(descGroup);

    // Meta Description
    const metaGroup = el('div', 'cc-seo-output-group');
    const metaHeader = el('div', 'cc-seo-output-header');
    const metaTitle = el('span', 'cc-seo-output-title', `Meta Description (${result.metaDescription.length} / 160 chars)`);
    const metaCopy = el('button', 'cc-seo-output-copy', 'Copy');
    metaCopy.addEventListener('click', () => {
      navigator.clipboard.writeText(result.metaDescription).then(() => {
        metaCopy.textContent = 'Copied!';
        setTimeout(() => { metaCopy.textContent = 'Copy'; }, 1500);
      });
    });
    append(metaHeader, metaTitle, metaCopy);
    const metaText = el('pre', 'cc-seo-output-text', result.metaDescription);
    append(metaGroup, metaHeader, metaText);
    container.appendChild(metaGroup);

    // Meta Keywords
    const keysGroup = el('div', 'cc-seo-output-group');
    const keysHeader = el('div', 'cc-seo-output-header');
    const keysTitle = el('span', 'cc-seo-output-title', 'Meta Keywords');
    const keysCopy = el('button', 'cc-seo-output-copy', 'Copy');
    keysCopy.addEventListener('click', () => {
      navigator.clipboard.writeText(result.metaKeywords).then(() => {
        keysCopy.textContent = 'Copied!';
        setTimeout(() => { keysCopy.textContent = 'Copy'; }, 1500);
      });
    });
    append(keysHeader, keysTitle, keysCopy);
    const keysText = el('pre', 'cc-seo-output-text', result.metaKeywords);
    append(keysGroup, keysHeader, keysText);
    container.appendChild(keysGroup);
  }

  function runCheck() {
    const pasteInputEl  = document.getElementById('cc-paste-input');
    const resultsDiv    = document.getElementById('cc-results');
    const checkBtn      = document.getElementById('cc-check-btn');
    const pasteRaw      = (pasteInputEl ? pasteInputEl.value : '').trim();

    if (!pasteRaw) {
      const pasteAreaEl = document.getElementById('cc-paste-area');
      if (pasteAreaEl) flashInput(pasteAreaEl);
      showError(resultsDiv, 'Please paste a citation list.');
      return;
    }

    // Spinner
    checkBtn.textContent = '';
    checkBtn.appendChild(el('span', 'cc-spinner'));
    checkBtn.disabled = true;

    setTimeout(() => {
      try {
        let dealerCells = [];
        let sourceNote  = '';
        let warnNote    = '';

        // ── Strategy 1: Pasted text ──────────────────────────────
        if (pasteRaw) {
          dealerCells = pasteRaw
            .split(/[\n\r,;\t]+/)
            .map(s => s.trim())
            .filter(s => s.length > 1);           // ignore single chars / blanks

          // Edge: suspicious volume → cap and warn
          if (dealerCells.length > 500) {
            warnNote = 'Only first 500 of ' + dealerCells.length + ' entries processed.';
            dealerCells = dealerCells.slice(0, 500);
          }
          sourceNote = 'Parsed ' + dealerCells.length + ' entries from pasted text';
        }

        // ── Edge: deduplicate before matching ────────────────────
        dealerCells = [...new Set(dealerCells)];

        // ── Matching ─────────────────────────────────────────────
        const matched   = new Map();
        const unmatched = [];

        for (const cell of dealerCells) {
          try {
            const m = findMatch(cell);
            if (m) {
              if (!matched.has(m.name)) matched.set(m.name, { entry: m, cellValue: cell });
            } else {
              // Only surface unmatched if it looks like a real URL/name (not pure numbers)
              if (cell.length > 3 && !/^\d+$/.test(cell)) unmatched.push(cell);
            }
          } catch (_) { /* skip malformed cell silently */ }
        }

        const doneEntries    = [...matched.values()];
        const missingEntries = MASTER_DIRECTORIES.filter(d => !matched.has(d.name));

        // ── Edge: zero matches with real input → hint user ───────
        if (doneEntries.length === 0 && dealerCells.length > 0 && !warnNote) {
          warnNote = 'No directories matched. Ensure the list contains domain names (e.g. yelp.com, bbb.org).';
        }

        const fullNote = sourceNote + (warnNote ? ' — ' + warnNote : '');
        renderResults(resultsDiv, doneEntries, missingEntries, unmatched, fullNote);
      } catch (err) {
        showError(resultsDiv, 'Unexpected error: ' + (err && err.message ? err.message : String(err)));
      } finally {
        checkBtn.textContent = 'Check';
        checkBtn.disabled = false;
      }
    }, 80);
  }

  function showError(container, msg) {
    if (!container) return;
    container.style.display = 'block';
    while (container.firstChild) container.removeChild(container.firstChild);
    container.appendChild(el('div', 'cc-error-box', msg));
  }

  function flashInput(inputEl) {
    inputEl.style.borderColor = '#ff7a00';
    inputEl.style.boxShadow   = '0 0 0 2px rgba(255,122,0,0.25)';
    setTimeout(() => { inputEl.style.borderColor = ''; inputEl.style.boxShadow = ''; }, 1200);
  }
  function mkStat(cls, num, label) {
    const stat = el('div', 'cc-stat ' + cls);
    append(stat, el('div', 'cc-stat-num', String(num)), el('div', 'cc-stat-label', label));
    return stat;
  }

  function mkCitationItem(cls, name, url, extraText, extraTitle) {
    const li   = el('li', 'cc-item ' + cls);
    const dot  = el('span', 'cc-dot');
    const nm   = el('span', 'cc-name', name);
    append(li, dot, nm);
    if (url) {
      const a = el('a', 'cc-link', extractDomain(url) || url);
      a.href   = url;
      a.target = '_blank';
      a.title  = url;
      li.appendChild(a);
    } else {
      const s = el('span', 'cc-link', 'no url');
      s.style.opacity = '0.3';
      li.appendChild(s);
    }
    if (extraText) {
      const via = el('span', 'cc-cell-match', 'via: ' + extraText.substring(0, 25));
      if (extraTitle) via.title = 'Matched from: ' + extraTitle;
      li.appendChild(via);
    }
    return li;
  }

  function renderResults(container, done, missing, unmatched, sourceNote) {
    container.style.display = 'block';

    // Clear previous results safely
    while (container.firstChild) container.removeChild(container.firstChild);

    const total = MASTER_DIRECTORIES.length;

    // ── Stats row ──────────────────────────────────────────
    const stats = el('div', 'cc-stats');
    append(stats,
      mkStat('done',  done.length,    'Listed'),
      mkStat('miss',  missing.length, 'Missing'),
      mkStat('total', total,          'Total')
    );
    container.appendChild(stats);

    // ── Source note ─────────────────────────────────────────
    const noteDiv = el('div', 'cc-source-note', sourceNote);
    container.appendChild(noteDiv);

    // ── Missing section ─────────────────────────────────────
    if (missing.length > 0) {
      container.appendChild(el('div', 'cc-section-title', 'Not Listed (' + missing.length + ')'));
      const ul = el('ul', 'cc-list');
      for (const entry of missing) ul.appendChild(mkCitationItem('miss', entry.name, entry.url));
      container.appendChild(ul);
    }

    // ── Done section ────────────────────────────────────────
    if (done.length > 0) {
      container.appendChild(el('div', 'cc-section-title', 'Already Listed (' + done.length + ')'));
      const ul = el('ul', 'cc-list');
      for (const { entry, cellValue } of done) {
        const via = (cellValue !== entry.name) ? cellValue : null;
        ul.appendChild(mkCitationItem('done', entry.name, entry.url, via, via));
      }
      container.appendChild(ul);
    }

    // ── Unmatched section ───────────────────────────────────
    if (unmatched.length > 0) {
      const box = el('div'); box.id = 'cc-unmatched';
      const ttl = el('div', 'cc-um-title', 'Unmatched — ' + unmatched.length + " cells didn't match any directory");
      const uli = el('ul');
      for (const u of unmatched.slice(0, 20)) uli.appendChild(el('li', '', u));
      if (unmatched.length > 20) uli.appendChild(el('li', '', '…and ' + (unmatched.length - 20) + ' more'));
      append(box, ttl, uli);
      container.appendChild(box);
    }

    // ── Copy missing button ─────────────────────────────────
    const missingText = missing.map(d => d.name).join('\n');
    const copyBtn = el('button', 'cc-export-btn', 'Copy Missing List to Clipboard');
    copyBtn.id = 'cc-copy-missing';
    copyBtn.addEventListener('click', () => {
      navigator.clipboard.writeText(missingText).then(() => {
        copyBtn.textContent = 'Copied!';
        setTimeout(() => { copyBtn.textContent = 'Copy Missing List to Clipboard'; }, 2000);
      }).catch(() => { window.prompt('Copy the missing list:', missingText); });
    });
    container.appendChild(copyBtn);
  }

  function escHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ─────────────────────────────────────────────────────────────
  //  INJECT STYLES  (works even if GM_addStyle is unavailable)
  // ─────────────────────────────────────────────────────────────
  function injectStyles() {
    if (document.getElementById('cc-styles')) return;
    const style = document.createElement('style');
    style.id = 'cc-styles';
    style.textContent = `
      /* Citation Sync – Orange and White Theme UI v3.1 */

      /* ── FAB trigger button ── */
      #cc-fab {
        position: fixed !important;
        bottom: 24px !important;
        right: 24px !important;
        z-index: 2147483647 !important;
        width: 50px !important;
        height: 50px !important;
        border-radius: 50% !important;
        background: #ffffff !important;
        border: 2px solid #ff7a00 !important;
        cursor: pointer !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-shadow: 0 4px 12px rgba(255,122,0,0.25) !important;
        transition: transform .15s, box-shadow .15s, border-color .15s !important;
        user-select: none !important;
        pointer-events: all !important;
        visibility: visible !important;
        opacity: 1 !important;
        padding: 4px !important;
        overflow: hidden !important;
        box-sizing: border-box !important;
      }
      #cc-fab:hover {
        transform: scale(1.08) !important;
        box-shadow: 0 6px 18px rgba(255,122,0,0.38) !important;
        border-color: #e06c00 !important;
      }
      #cc-fab img {
        width: 100% !important;
        height: 100% !important;
        object-fit: contain !important;
        display: block !important;
        border-radius: 50% !important;
      }

      /* ── Main panel ── */
      #cc-panel {
        position: fixed !important;
        bottom: 82px !important;
        right: 24px !important;
        z-index: 2147483646 !important;
        width: 360px !important;
        max-height: 78vh !important;
        background: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 12px !important;
        box-shadow: 0 10px 40px rgba(0,0,0,0.12) !important;
        font-family: 'Segoe UI', system-ui, -apple-system, sans-serif !important;
        color: #334155 !important;
        overflow: hidden !important;
        flex-direction: column !important;
        pointer-events: all !important;
        visibility: visible !important;
        opacity: 1 !important;
      }

      /* ── Header ── */
      #cc-header {
        background: #ff7a00 !important;
        padding: 12px 16px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: space-between !important;
        flex-shrink: 0 !important;
      }
      #cc-header h2 {
        margin: 0;
        font-size: 14px;
        font-weight: 700;
        color: #ffffff;
        letter-spacing: .02em;
        text-transform: uppercase;
      }
      #cc-header span {
        font-size: 10px;
        color: rgba(255,255,255,0.85);
        display: block;
        margin-top: 2px;
      }
      #cc-close {
        background: rgba(255,255,255,0.2);
        border: 1px solid rgba(255,255,255,0.3);
        color: #ffffff;
        border-radius: 6px;
        width: 26px;
        height: 26px;
        cursor: pointer;
        font-size: 18px;
        line-height: 1;
        display: flex;
        align-items: center;
        justify-content: center;
        transition: background .15s;
        flex-shrink: 0;
      }
      #cc-close:hover { background: rgba(255,255,255,0.35); }

      /* ── Body ── */
      #cc-body { padding: 14px; overflow-y: auto; flex: 1; background: #f8fafc; }
      #cc-body::-webkit-scrollbar { width: 6px; }
      #cc-body::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }

      /* ── Tabs ── */
      .cc-tabs {
        display: flex;
        border-bottom: 2px solid #e2e8f0;
        margin-bottom: 12px;
        background: #ffffff;
      }
      .cc-tab {
        flex: 1;
        text-align: center;
        padding: 10px;
        font-size: 13px;
        font-weight: 600;
        color: #64748b;
        cursor: pointer;
        transition: color .15s, border-color .15s;
        border-bottom: 2px solid transparent;
        margin-bottom: -2px;
        user-select: none;
      }
      .cc-tab:hover { color: #ff7a00; }
      .cc-tab.active {
        color: #ff7a00;
        border-bottom-color: #ff7a00;
      }

      /* ── Search row ── */
      #cc-search-row { display: flex; gap: 8px; margin-bottom: 12px; }
      #cc-check-btn, #cc-seo-gen-btn {
        background: #ff7a00;
        border: 1px solid #e06c00;
        border-radius: 6px;
        color: #ffffff;
        font-size: 13px;
        font-weight: 600;
        padding: 10px 20px;
        cursor: pointer;
        white-space: nowrap;
        font-family: inherit;
        transition: background .15s, border-color .15s, transform .1s, box-shadow .15s;
        box-shadow: 0 2px 4px rgba(255,122,0,0.1);
      }
      #cc-check-btn:hover, #cc-seo-gen-btn:hover {
        background: #f97316;
        border-color: #ea580c;
        box-shadow: 0 4px 8px rgba(255,122,0,0.2);
      }
      #cc-check-btn:active, #cc-seo-gen-btn:active {
        transform: scale(0.98);
        box-shadow: 0 1px 2px rgba(255,122,0,0.1);
      }
      #cc-check-btn:disabled, #cc-seo-gen-btn:disabled {
        background: #cbd5e1;
        border-color: #cbd5e1;
        color: #94a3b8;
        cursor: not-allowed;
        box-shadow: none;
        transform: none;
      }

      /* ── Paste area ── */
      #cc-paste-area {
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 10px 12px;
        margin-bottom: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
      }
      #cc-paste-area label {
        font-size: 10px;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: .06em;
        display: block;
        margin-bottom: 6px;
        font-weight: 600;
      }
      #cc-paste-input {
        width: 100%;
        background: transparent;
        border: none;
        color: #334155;
        font-size: 12px;
        resize: vertical;
        outline: none;
        min-height: 60px;
        font-family: 'Consolas', 'Cascadia Code', monospace;
      }
      #cc-paste-input::placeholder { color: #cbd5e1; }

      /* ── SEO Generator Input Styles ── */
      .cc-input-group { margin-bottom: 10px; }
      .cc-input-label {
        font-size: 10px;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: .06em;
        display: block;
        margin-bottom: 4px;
        font-weight: 600;
      }
      .cc-text-input {
        width: 100%;
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 8px 12px;
        color: #1e293b;
        font-size: 13px;
        outline: none;
        transition: border-color .15s, box-shadow .15s;
        font-family: inherit;
        box-sizing: border-box;
      }
      .cc-text-input:focus { border-color: #ff7a00; box-shadow: 0 0 0 3px rgba(255,122,0,0.15); }
      /* cc-subtle-btn and cc-scrape-row removed */
      .cc-rules-box {
        background: #fffbeb;
        border: 1px solid #fef3c7;
        border-radius: 8px;
        padding: 10px;
        margin-bottom: 14px;
        font-size: 11px;
        color: #b45309;
      }
      .cc-rules-title { font-weight: 700; margin-bottom: 4px; text-transform: uppercase; letter-spacing: 0.02em; }
      .cc-rules-box ul { margin: 0; padding-left: 14px; }
      .cc-rules-box li { margin-bottom: 2px; }

      /* ── SEO Generator Output Styles ── */
      .cc-seo-output-group {
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 10px;
        margin-bottom: 10px;
        position: relative;
      }
      .cc-seo-output-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 6px;
      }
      .cc-seo-output-title { font-size: 11px; font-weight: 700; color: #475569; text-transform: uppercase; }
      .cc-seo-output-copy {
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        color: #475569;
        font-size: 11px;
        padding: 4px 10px;
        cursor: pointer;
        font-family: inherit;
        font-weight: 600;
        transition: background .15s, border-color .15s;
        box-shadow: 0 1px 2px rgba(0,0,0,0.02);
      }
      .cc-seo-output-copy:hover {
        background: #f8fafc;
        border-color: #94a3b8;
        color: #1e293b;
      }
      .cc-seo-output-copy:active {
        background: #f1f5f9;
      }
      .cc-seo-output-text {
        font-size: 12px;
        color: #1e293b;
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 6px;
        padding: 8px;
        margin: 0;
        white-space: pre-wrap;
        word-break: break-word;
        font-family: inherit;
        max-height: 80px;
        overflow-y: auto;
      }

      /* ── Stats row ── */
      .cc-stats { display: flex; gap: 8px; margin-bottom: 14px; }
      .cc-stat {
        flex: 1;
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 10px 8px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
      }
      .cc-stat .cc-stat-num { font-size: 20px; font-weight: 700; line-height: 1; }
      .cc-stat .cc-stat-label { font-size: 10px; color: #64748b; margin-top: 4px; font-weight: 600; text-transform: uppercase; }
      .cc-stat.done  .cc-stat-num { color: #10b981; }
      .cc-stat.miss  .cc-stat-num { color: #ff7a00; }
      .cc-stat.total .cc-stat-num { color: #6366f1; }

      /* ── Section titles ── */
      .cc-section-title {
        font-size: 10px;
        text-transform: uppercase;
        letter-spacing: .08em;
        color: #64748b;
        margin: 14px 0 6px;
        font-weight: 700;
        border-bottom: 1px solid #e2e8f0;
        padding-bottom: 6px;
      }

      /* ── List items ── */
      .cc-list { list-style: none; margin: 0 0 10px; padding: 0; }
      .cc-item {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 6px 8px;
        border-radius: 6px;
        font-size: 13px;
        transition: background .1s;
      }
      .cc-item:hover { background: #f1f5f9; }
      .cc-item .cc-dot { width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }
      .cc-item.done .cc-dot { background: #10b981; }
      .cc-item.miss .cc-dot { background: #ff7a00; }
      .cc-item .cc-name { flex: 1; color: #334155; }
      .cc-item .cc-link {
        font-size: 11px;
        color: #64748b;
        text-decoration: none;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        max-width: 110px;
      }
      .cc-item .cc-link:hover { color: #ff7a00; text-decoration: underline; }
      .cc-item .cc-cell-match {
        font-size: 10px;
        color: #94a3b8;
        font-style: italic;
        max-width: 90px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }

      /* ── Unmatched box ── */
      #cc-unmatched {
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 10px 12px;
        margin-top: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
      }
      #cc-unmatched .cc-um-title {
        font-size: 11px;
        color: #d97706;
        margin-bottom: 6px;
        font-weight: 600;
      }
      #cc-unmatched ul { margin: 0; padding: 0 0 0 16px; font-size: 12px; color: #475569; }
      #cc-unmatched ul li { margin-bottom: 3px; }

      /* ── Error box ── */
      .cc-error-box {
        background: #fff5f5;
        border: 1px solid #feb2b2;
        border-radius: 8px;
        padding: 10px 12px;
        font-size: 13px;
        color: #c53030;
        margin-top: 6px;
      }

      /* ── Hint ── */
      #cc-hint {
        color: #64748b;
        font-size: 12px;
        padding: 8px 0 4px;
        line-height: 1.6;
        text-align: center;
      }

      /* ── Copy button ── */
      .cc-export-btn {
        display: flex;
        align-items: center;
        justify-content: center;
        width: 100%;
        background: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        color: #334155;
        font-size: 13px;
        font-weight: 600;
        padding: 11px;
        cursor: pointer;
        margin-top: 12px;
        font-family: inherit;
        transition: background .15s, border-color .15s, transform .1s;
        box-shadow: 0 1px 2px rgba(0,0,0,0.02);
      }
      .cc-export-btn:hover {
        background: #f8fafc;
        border-color: #94a3b8;
      }
      .cc-export-btn:active {
        transform: scale(0.99);
      }

      /* ── Source note ── */
      .cc-source-note {
        font-size: 11px;
        color: #475569;
        margin-bottom: 12px;
        padding: 6px 10px;
        background: #ffffff;
        border-left: 3px solid #ff7a00;
        border-radius: 0 6px 6px 0;
        box-shadow: 0 1px 2px rgba(0,0,0,0.01);
      }

      /* ── Spinner ── */
      .cc-spinner {
        display: inline-block;
        width: 14px; height: 14px;
        border: 2px solid rgba(255,255,255,.25);
        border-top-color: #fff;
        border-radius: 50%;
        animation: cc-spin .6s linear infinite;
        vertical-align: middle;
      }
      @keyframes cc-spin { to { transform: rotate(360deg); } }
    `;
    (document.head || document.documentElement).appendChild(style);
  }

  // ─────────────────────────────────────────────────────────────
  //  INIT  –  robust retry until body is available
  // ─────────────────────────────────────────────────────────────
  function init() {
    if (document.getElementById('cc-fab')) return; // already injected
    if (!document.body) return;                    // body not ready yet

    try {
      injectStyles();
      buildUI();
      console.log('[Citation Sync v3.0] Loaded on', location.hostname,
                  '| Directories:', MASTER_DIRECTORIES.length);
    } catch (err) {
      console.error('[Citation Sync] Init error:', err);
    }
  }

  // ── Retry loop: try every 500ms for up to 15 seconds ──────────
  // Handles SPAs (Excel Online, Google Sheets) that build their DOM
  // asynchronously long after document-idle fires.
  let _retryCount = 0;
  const MAX_RETRIES = 30; // 30 × 500ms = 15 seconds

  function tryInit() {
    if (document.getElementById('cc-fab')) return; // already done
    if (_retryCount >= MAX_RETRIES) {
      console.warn('[Citation Sync] Gave up after 15s. Page may block script injection.');
      return;
    }
    _retryCount++;
    init();
    if (!document.getElementById('cc-fab')) {
      setTimeout(tryInit, 500);
    }
  }

  // Also re-inject if the SPA navigates and wipes the DOM
  const _observer = new MutationObserver(() => {
    if (!document.getElementById('cc-fab') && document.body) {
      init();
    }
  });

  function startObserver() {
    if (document.body) {
      _observer.observe(document.body, { childList: true, subtree: false });
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => { startObserver(); tryInit(); });
  } else {
    startObserver();
    tryInit();
  }

})();
