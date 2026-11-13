({ customSelector = "" } = {}) => {
  const clean = value => String(value || "").replace(/\s+/g, " ").trim().slice(0, 240);
  const name = el => clean(el.getAttribute("aria-label") || el.getAttribute("title") || el.textContent || el.querySelector("img")?.alt);
  const roots = [document];
  for (let i = 0; i < roots.length; i++) {
    for (const el of roots[i].querySelectorAll("*")) if (el.shadowRoot) roots.push(el.shadowRoot);
  }
  // Standard menu markup first, then names older bank/government sites use (#menu, #topMenu, header lists, *gnb*, *nav*).
  const navSelector = customSelector || "nav,[role=navigation],[role=menubar],[role=menu],.gnb,#gnb,.lnb,#lnb,.snb,#snb,[class*=lnb],[class*=snb],.sitemap,#sitemap,[data-menu-root]";
  const legacySelector = "[id*=gnb i],[class*=gnb i],#menu,.menu,[id$=menu i],[id*=mainmenu i],[id*=topmenu i],[class*=mainmenu i],[class*=main-menu i],[id*=navi i],[class*=navi i],header ul,#header ul,.header ul";
  const menuArea = el => {
    const nav = el.closest(navSelector);
    if (nav || customSelector) return nav;
    let top = null;
    for (let c = el.closest(legacySelector); c; c = c.parentElement?.closest(legacySelector)) top = c;   // outermost legacy area
    return top && top !== document.body ? top : null;
  };
  const crumbSelector = '.breadcrumb,.breadcrumbs,[aria-label*="breadcrumb" i],.location-path,[data-breadcrumb]';
  const isHome = value => /^(home|홈|메인|첫 화면)$/i.test(value);
  // Side menus (LNB/SNB) usually show their parent as a heading instead of a nested list.
  const sideSelector = '.lnb,#lnb,.snb,#snb,[class*="lnb"],[class*="snb"],[class*="side-menu"],[class*="sidemenu"],aside';
  const commonSelector = 'footer,#footer,.footer,[class*="footer"],.util,#util,[class*="util"],.top-menu,.topmenu,[class*="family"]';
  const hiddenHeading = el => /(?:^|\s)(?:blind|sr-only|screen_out|hide|hidden|skip|a11y|ir)(?:\s|$)/i.test(el.className || "")
    || /^(?:주\s*메뉴|메인\s*메뉴|전체\s*메뉴|메뉴|navigation|main menu|menu|global navigation)$/i.test(clean(el.textContent));
  const sideTitle = nav => {
    const side = nav.closest(sideSelector);
    if (!side) return "";
    const heading = side.querySelector("h1,h2,h3,h4,.title,.tit,[class*='title'],[class*='tit']");
    if (!heading || heading.closest("li") || hiddenHeading(heading)) return "";
    return clean(heading.textContent);
  };
  let breadcrumbs = [];
  const crumb = document.querySelector(crumbSelector);
  if (crumb) {
    const elements = crumb.querySelectorAll("li").length ? crumb.querySelectorAll("li") : crumb.querySelectorAll("a,span");
    breadcrumbs = [...elements].map(name).filter(x => x && !isHome(x));
    breadcrumbs = breadcrumbs.filter((x, i) => x !== breadcrumbs[i - 1]);
  }
  const records = [];
  let unnamed = 0, seq = 0;
  for (const root of roots) {
    // Validate the optional selector even on a page with no links.
    root.querySelectorAll(navSelector);
    for (const el of root.querySelectorAll('a,button,[role="menuitem"],[role="treeitem"],[data-href],[data-url]')) {
      if (el.closest(crumbSelector)) continue;
      let nav = menuArea(el);
      if (nav && !customSelector && nav.closest(commonSelector)) nav = null;   // footer / utility lists are not menus
      const raw = el.getAttribute("href") || el.getAttribute("data-href") || el.getAttribute("data-url") || "";
      if (!nav && !raw) continue;
      // In-page anchors (#content, #top, skip links) are not menus. SPA routes (#/x, #!/x) are.
      if (/^#[^/!]/.test(raw) && !el.getAttribute("onclick")) continue;
      const label = name(el);
      if (!label) { unnamed++; continue; }
      const leaf = el.closest('li,[role="treeitem"]');
      const controlled = el.getAttribute("aria-controls");
      const target = controlled ? root.getElementById?.(controlled) : null;
      const children = (leaf && [...leaf.children].some(x => x.matches('ul,ol,[role="menu"],[role="group"]'))) || Boolean(target?.querySelector("a,button,[role=menuitem]"));
      const parentLabel = item => {
        for (const child of item.children) {
          if (child.matches("a,button,span,[role=menuitem]")) {
            const copy = child.cloneNode(true);
            copy.querySelectorAll("ul,ol,[role=menu],[role=group]").forEach(n => n.remove());
            const value = name(copy);
            if (value) return value;
          }
        }
        return clean([...item.childNodes].filter(n => n.nodeType === Node.TEXT_NODE).map(n => n.textContent).join(" "));
      };
      const ancestors = [];
      let cursor = leaf?.parentElement?.closest('li,[role="treeitem"]');
      while (cursor && (!nav || nav.contains(cursor)) && ancestors.length < 20) {
        const text = parentLabel(cursor);
        if (text) ancestors.unshift(text);
        cursor = cursor.parentElement?.closest('li,[role="treeitem"]');
      }
      if (nav?.id && ancestors.length === 0) {
        const owner = [...root.querySelectorAll("[aria-controls]")].find(x => x.getAttribute("aria-controls") === nav.id);
        if (owner && owner !== el) { const text = name(owner); if (text && text !== label) ancestors.push(text); }
      }
      if (nav && ancestors.length === 0) {
        const parent = sideTitle(nav);
        if (parent && parent !== label) ancestors.push(parent);
      }
      let destination = raw;
      if (!raw || raw === "#" || /^javascript:/i.test(raw)) {
        const script = (el.getAttribute("onclick") || "") + " " + raw;
        const match = script.match(/(?:location(?:\.href)?\s*=\s*|(?:window\.)?open\s*\(\s*)['"]([^'"]+)['"]/);
        destination = match ? match[1] : "";
      }
      if (/^(mailto|tel|sms):/i.test(destination)) continue;
      let url = "";
      try { if (destination) { const u = new URL(destination, document.baseURI); if (/^https?:$/.test(u.protocol)) url = u.href; } } catch {}
      const common = !nav && el.closest(commonSelector);
      const path = nav || common ? [...ancestors, label] : [...breadcrumbs, label];
      const evidence = nav ? "메뉴 DOM" : common ? "공통 링크" : breadcrumbs.length ? "본문 링크·현재 경로" : "일반 링크";
      const mc = String(seq++);
      el.setAttribute("data-mc-id", mc);
      records.push({ path, url, label, kind: !url && children ? "group" : !url ? "dynamic" : "link", evidence, hierarchy_review: !nav, mc });
    }
  }
  // A login form, not just any password field (an account PIN on a transfer page is not a login screen):
  // a visible password field with an ID-like field next to it and login words around it.
  const visible = el => Boolean(el && (el.offsetWidth || el.offsetHeight || el.getClientRects().length));
  const loginWords = /(로그인|log\s*-?\s*in|sign\s*-?\s*in|아이디|user\s*id|사용자\s*id|forgot|비밀번호\s*찾기|아이디\s*찾기)/i;
  const isLoginForm = field => {
    let box = field.closest("form") || field.parentElement;
    for (let i = 0; i < 3 && box && !box.querySelector('input:not([type]),input[type="text"],input[type="email"],input[type="tel"]'); i++) box = box.parentElement;
    if (!box) return false;
    const idField = [...box.querySelectorAll('input:not([type]),input[type="text"],input[type="email"],input[type="tel"]')].some(visible);
    return idField && loginWords.test(clean(box.innerText || box.textContent || "") + " " + document.title);
  };
  const pw = [...document.querySelectorAll('input[type="password"]')].find(f => visible(f) && isLoginForm(f));
  // Functions of the login box (find ID, reset password, register...): menus of their own for a diagnosis.
  if (pw) {
    // The login box: the form, widened a few levels while it stays small and holds no menu area.
    const holdsMenu = el => el.matches(navSelector + "," + legacySelector) || el.querySelector(navSelector + "," + legacySelector);
    let box = pw.closest("form") || pw.parentElement;
    for (let i = 0; i < 4 && box?.parentElement && box.parentElement !== document.body
         && box.parentElement.querySelectorAll("a,button").length <= 12 && !holdsMenu(box.parentElement); i++) box = box.parentElement;
    const loginNav = records.find(r => r.evidence === "메뉴 DOM" && r.path.length === 1 && /^(log\s*in|sign\s*in|로그인)$/i.test(r.label));
    const top = loginNav ? loginNav.label : (/^ko/i.test(document.documentElement.lang || "ko") ? "로그인" : "Login");
    for (const el of box ? box.querySelectorAll('a,button,[role="button"],input[type="button"]') : []) {
      if (el.matches('[type="submit"]') || menuArea(el)) continue;
      const label = clean(el.value || name(el));
      if (!label || label.length > 60 || /^(log\s*in|sign\s*in|로그인|ok|확인|submit|cancel|취소|close|닫기)$/i.test(label)) continue;
      const raw = el.getAttribute("href") || "";
      let url = "";
      try { if (raw && !/^(javascript:|#)/i.test(raw)) { const u = new URL(raw, document.baseURI); if (/^https?:$/.test(u.protocol)) url = u.href; } } catch {}
      const mc = String(seq++);
      el.setAttribute("data-mc-id", mc);
      records.push({ path: [top, label], url, label, kind: url ? "link" : "dynamic", evidence: "로그인 화면", hierarchy_review: false, mc });
    }
  }
  return {
    records, breadcrumbs, title: document.title,
    password: Boolean(pw),
    loginText: /(?:로그인(?:이|을|이\s*다시)?\s*(?:필요|해\s*주세요|하십시오)|로그인\s*후\s*(?:이용|접근)|sign\s*in\s*required|please\s*(?:log|sign)\s*in)/i.test(document.body?.innerText || ""),
    unnamed,
  };
}
