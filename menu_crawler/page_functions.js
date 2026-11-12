() => {
  // Tabs and choices inside the page content (e.g. "User ID issued at the branch" on a find-ID page).
  const clean = v => String(v || "").replace(/\s+/g, " ").trim();
  const out = [], seen = new Set();
  const skip = el => el.closest('nav,header,#header,footer,#footer,[role=navigation],#gnb,.gnb,#lnb,.lnb,[class*=util],[data-menu-crawler-hidden]');
  const push = t => { t = clean(t); if (t && t.length <= 80 && !seen.has(t)) { seen.add(t); out.push(t); } };
  const tabs = document.querySelectorAll('[role=tab],[class*=tab i] > li > a,[class*=tab i] > li > button,[class*=tab i] > a,[class*=tab i] > button');
  for (const el of tabs) if (!skip(el)) push(el.textContent);
  if (!out.length) {
    for (const r of document.querySelectorAll('input[type=radio]')) {
      const group = r.name ? document.querySelectorAll(`input[type=radio][name="${CSS.escape(r.name)}"]`) : [r];
      if (group.length < 2 || group.length > 6) continue;
      const label = (r.id && document.querySelector(`label[for="${CSS.escape(r.id)}"]`)) || r.closest("label");
      const text = clean(label?.textContent);
      // a choice that names a way of doing something, not a short field value like "Male"
      if (label && !skip(label) && (text.split(" ").length >= 3 || /[가-힣]{6,}/.test(text))) push(text);
    }
  }
  return out.slice(0, 20);
}
