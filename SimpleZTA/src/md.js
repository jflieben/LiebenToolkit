(() => {
  // Markdown utilities shared by the official report export (module-shaped TestDescription /
  // TestResult fields) and the result detail modal. The bundled module ships one markdown file
  // per control at powershell/tests/Test-Assessment.<id>.md, split by an optional
  // `<!--- Results --->` marker into a description part and a result template part. The result
  // template may contain a %TestResult% placeholder (same contract as Add-ZtTestResultDetail).

  const RESULTS_MARKER = /<!---?\s*Results\s*---?>/i;
  const mdCache = new Map();

  function looksLikeHtmlDocument(text) {
    const sample = `${text || ''}`.trimStart().slice(0, 512).toLowerCase();
    return sample.startsWith('<!doctype html') || sample.startsWith('<html');
  }

  function escapeHtml(value) {
    return `${value ?? ''}`
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function safeUrl(url) {
    const u = `${url || ''}`.trim();
    return /^https?:\/\//i.test(u) ? u : '';
  }

  // Inline markdown: code spans, bold, italics, links. Input must already be HTML-escaped.
  function renderInline(text) {
    let out = text;
    out = out.replace(/`([^`]+)`/g, (_, code) => `<code>${code}</code>`);
    out = out.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    out = out.replace(/(^|[\s(])\*([^*\s][^*]*)\*/g, '$1<em>$2</em>');
    out = out.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (m, label, url) => {
      const href = safeUrl(url.replace(/&amp;/g, '&'));
      if (!href) return label;
      return `<a href="${escapeHtml(href)}" target="_blank" rel="noopener noreferrer">${label}</a>`;
    });
    return out;
  }

  function isTableDivider(line) {
    return /^\s*\|?[\s:|-]+\|[\s:|-]*$/.test(line) && line.includes('-');
  }

  function splitRow(line) {
    let l = line.trim();
    if (l.startsWith('|')) l = l.slice(1);
    if (l.endsWith('|')) l = l.slice(0, -1);
    return l.split('|').map(c => c.trim());
  }

  // Small, dependency-free markdown-to-HTML renderer. All raw input is HTML-escaped before
  // any transformation, so untrusted markdown cannot inject markup.
  function mdToHtml(md) {
    const lines = `${md || ''}`.replace(/\r\n/g, '\n').split('\n');
    const out = [];
    let i = 0;
    let paragraph = [];
    let listStack = null; // { type: 'ul'|'ol', items: [] }

    function flushParagraph() {
      if (paragraph.length) {
        out.push(`<p>${renderInline(paragraph.join(' '))}</p>`);
        paragraph = [];
      }
    }

    function flushList() {
      if (listStack) {
        out.push(`<${listStack.type}>${listStack.items.map(x => `<li>${x}</li>`).join('')}</${listStack.type}>`);
        listStack = null;
      }
    }

    while (i < lines.length) {
      const raw = lines[i];
      const line = escapeHtml(raw);

      // Fenced code block
      if (/^\s*```/.test(raw)) {
        flushParagraph(); flushList();
        const buf = [];
        i += 1;
        while (i < lines.length && !/^\s*```/.test(lines[i])) {
          buf.push(escapeHtml(lines[i]));
          i += 1;
        }
        i += 1; // skip closing fence
        out.push(`<pre><code>${buf.join('\n')}</code></pre>`);
        continue;
      }

      // Table
      if (raw.includes('|') && i + 1 < lines.length && isTableDivider(lines[i + 1])) {
        flushParagraph(); flushList();
        const headers = splitRow(line).map(renderInline);
        i += 2;
        const rows = [];
        while (i < lines.length && lines[i].includes('|') && lines[i].trim() !== '') {
          rows.push(splitRow(escapeHtml(lines[i])).map(renderInline));
          i += 1;
        }
        out.push(`<div class="md-table-wrap"><table class="results-table"><thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.map(r => `<tr>${r.map(c => `<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`);
        continue;
      }

      // Heading
      const heading = /^(#{1,6})\s+(.*)$/.exec(raw);
      if (heading) {
        flushParagraph(); flushList();
        const level = Math.min(6, heading[1].length + 2); // demote: md h1 -> h3 inside the modal
        out.push(`<h${level}>${renderInline(escapeHtml(heading[2]))}</h${level}>`);
        i += 1;
        continue;
      }

      // List item (allow leading indent for nested items; render flat)
      const li = /^\s*([-*+]|\d+\.)\s+(.*)$/.exec(raw);
      if (li) {
        flushParagraph();
        const type = /^\d+\.$/.test(li[1]) ? 'ol' : 'ul';
        if (!listStack || listStack.type !== type) {
          flushList();
          listStack = { type, items: [] };
        }
        listStack.items.push(renderInline(escapeHtml(li[2])));
        i += 1;
        continue;
      }

      // Blank line
      if (raw.trim() === '') {
        flushParagraph(); flushList();
        i += 1;
        continue;
      }

      paragraph.push(line);
      i += 1;
    }

    flushParagraph(); flushList();
    return out.join('\n');
  }

  // Loads the module markdown for a test id. Returns { description, resultTemplate } or null.
  async function loadTestMd(id) {
    const key = `${id}`;
    if (mdCache.has(key)) return mdCache.get(key);
    let entry = null;
    try {
      const res = await fetch(`./powershell/tests/Test-Assessment.${key}.md`, { cache: 'force-cache' });
      if (res.ok) {
        const contentType = `${res.headers?.get?.('content-type') || ''}`.toLowerCase();
        const text = await res.text();
        // Some hosts rewrite unknown paths to an HTML page (often index.html) while still
        // returning 200. Guard against that so report remediation sections never contain
        // full page markup.
        if (!contentType.includes('text/html') && !looksLikeHtmlDocument(text)) {
          const parts = text.split(RESULTS_MARKER);
          entry = {
            description: (parts[0] || '').trim(),
            resultTemplate: (parts[1] || '').trim(),
          };
        }
      }
    } catch {
      entry = null;
    }
    mdCache.set(key, entry);
    return entry;
  }

  async function loadTestMdMany(ids, concurrency) {
    const queue = Array.from(new Set((ids || []).map(x => `${x}`)));
    const result = new Map();
    const workers = Array.from({ length: Math.max(1, concurrency || 8) }, async () => {
      while (queue.length) {
        const id = queue.shift();
        result.set(id, await loadTestMd(id));
      }
    });
    await Promise.all(workers);
    return result;
  }

  // Admin portal deep links per evidence object type (ported from Get-GraphObjectMarkdown).
  const PORTAL_LINKS = {
    AuthenticationMethod: 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/AuthenticationMethodsMenuBlade/~/AdminAuthMethods',
    AuthenticationStrength: 'https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade/~/AuthStrengths/fromNav/',
    AuthorizationPolicy: 'https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade/~/UserSettings/menuId/UserSettings',
    ConditionalAccess: 'https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/PolicyBlade/policyId/{0}',
    ConsentPolicy: 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ConsentPoliciesMenuBlade/~/UserSettings',
    Devices: 'https://entra.microsoft.com/#view/Microsoft_AAD_Devices/DeviceDetailsMenuBlade/~/Properties/objectId/{0}',
    DiagnosticSettings: 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/DiagnosticSettingsMenuBlade/~/General',
    Domains: 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/DomainsManagementMenuBlade/~/CustomDomainNames',
    Groups: 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/{0}',
    IdentityProtection: 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/IdentityProtectionMenuBlade/~/UsersAtRiskAlerts/fromNav/Identity',
    Users: 'https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/{0}',
    UserRole: 'https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/AdministrativeRole/userId/{0}',
    Applications: 'https://entra.microsoft.com/#view/Applications/ApplicationMenuBlade/~/Overview/appId/{0}',
  };

  function portalLink(type, id) {
    const template = PORTAL_LINKS[type];
    if (!template) return '';
    return template.replace('{0}', encodeURIComponent(`${id || ''}`));
  }

  // Detects evidence arrays whose items look like Graph directory objects and renders a
  // markdown list with portal deep links, mirroring Get-GraphObjectMarkdown output.
  function graphObjectsToMarkdown(items, type) {
    const rows = (items || []).filter(x => x && (x.displayName || x.id));
    if (!rows.length) return '';
    return rows.map(item => {
      const name = item.displayName || item.userPrincipalName || item.id;
      let suffix = '';
      if (type === 'ConditionalAccess') {
        if (item.state === 'disabled') suffix = ' (Disabled)';
        else if (item.state === 'enabledForReportingButNotEnforced') suffix = ' (Report-only)';
      }
      const link = portalLink(type, item.id);
      return link ? `- [${name}](${link})${suffix}` : `- ${name}${suffix}`;
    }).join('\n');
  }

  // Builds the module-style TestResult markdown for a web-app scan result.
  function buildResultMarkdown(row, resultTemplate) {
    const statusIcon = { Passed: '✅', Failed: '❌', Skipped: '🟦', Error: '❗', Investigate: '🔍' }[row.status] || '';
    const parts = [];
    parts.push(`${statusIcon} **${row.status}** — ${row.message || 'No summary recorded.'}`.trim());

    if (row.implementationLevel === 'partial' && row.implementationNotes) {
      parts.push([
        '',
        '> **Browser implementation note (SimpleZTA)**',
        `> Implemented: ${row.implementationNotes.implemented || '-'}`,
        `> Not implemented: ${row.implementationNotes.notImplemented || '-'}`,
      ].join('\n'));
    }

    if (row.evidence != null) {
      let evidenceText = '';
      try {
        evidenceText = typeof row.evidence === 'string' ? row.evidence : JSON.stringify(row.evidence, null, 2);
      } catch {
        evidenceText = `${row.evidence}`;
      }
      if (evidenceText.length > 6000) evidenceText = `${evidenceText.slice(0, 6000)}\n...truncated...`;
      // Fences inside evidence would break the block; indent-neutralize them.
      evidenceText = evidenceText.replace(/```/g, '`​``');
      parts.push(`\n**Evidence**\n\n\`\`\`json\n${evidenceText}\n\`\`\``);
    }

    const body = parts.join('\n');
    const template = `${resultTemplate || ''}`.trim();
    if (template && template.includes('%TestResult%')) {
      return template.replace(/%TestResult%/g, body);
    }
    return body;
  }

  window.Md = {
    mdToHtml,
    escapeHtml,
    loadTestMd,
    loadTestMdMany,
    portalLink,
    graphObjectsToMarkdown,
    buildResultMarkdown,
  };
})();
