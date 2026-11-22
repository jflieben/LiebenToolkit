// Google Analytics, the same snippet as lieben.nu, the blog and waarmerk.jsolve.nl: gtag.js with
// 'config' BEFORE 'consent' (consent first, or a hand-made beacon, is not counted by GA4).
// The toolkit allows no inline scripts, so this file loads gtag.js itself.
(() => {
  const MEASUREMENT_ID = 'G-20PYJ8S2G7';
  try {
    const host = location.hostname;
    if (host === 'localhost' || host === '127.0.0.1' || host === '' || host === '::1') return;
    // MSAL's sign-in popup loads the tool page too; that is not a page view.
    if (window.opener && String(window.name || '').startsWith('msal.')) return;
    window.dataLayer = window.dataLayer || [];
    function gtag() { window.dataLayer.push(arguments); }
    window.gtag = gtag;
    gtag('js', new Date());
    gtag('config', MEASUREMENT_ID, {
      anonymize_ip: true,
      storage: 'none',
      client_storage: 'none',
      // Only the path: query and hash can carry sign-in responses
      page_location: `${location.origin}${location.pathname}`,
    });
    gtag('consent', 'default', {
      ad_storage: 'denied',
      ad_user_data: 'denied',
      ad_personalization: 'denied',
      analytics_storage: 'denied',
    });
    const script = document.createElement('script');
    script.async = true;
    script.src = `https://www.googletagmanager.com/gtag/js?id=${MEASUREMENT_ID}`;
    document.head.append(script);
  } catch { /* never let statistics break a tool */ }
})();
