// Pure helpers for InactiveGuests - safe to unit-test.
(() => {
  const ZERO_DATE = '0001-01-01T00:00:00Z';

  /**
   * Pick the most recent of interactive vs non-interactive sign-in.
   * Returns a Date or null when the guest has never signed in.
   */
  function lastSignInDate(user) {
    const sa = user && user.signInActivity;
    if (!sa) return null;
    const candidates = [sa.lastSignInDateTime, sa.lastNonInteractiveSignInDateTime]
      .filter(v => v && v !== ZERO_DATE)
      .map(v => new Date(v))
      .filter(d => !isNaN(d.getTime()));
    if (!candidates.length) return null;
    return new Date(Math.max(...candidates.map(d => d.getTime())));
  }

  function daysBetween(a, b) {
    if (!a || !b) return null;
    return Math.floor((b.getTime() - a.getTime()) / 86400000);
  }

  /**
   * Enrich a guest record with computed fields used everywhere in the UI.
   *   lastSignIn           Date | null
   *   inactiveDays         number  - days since last sign-in OR since creation if never
   *   accountAgeDays       number  - days since createdDateTime
   *   neverSignedIn        bool
   *   redemptionState      'PendingAcceptance' | 'Accepted' | 'Unknown'
   */
  function enrich(user, now = new Date()) {
    const last = lastSignInDate(user);
    const created = user.createdDateTime ? new Date(user.createdDateTime) : null;
    const accountAgeDays = created ? daysBetween(created, now) : null;
    const inactiveDays = last ? daysBetween(last, now) : accountAgeDays;
    return {
      ...user,
      lastSignIn: last,
      inactiveDays,
      accountAgeDays,
      neverSignedIn: !last,
      redemptionState: user.externalUserState === 'PendingAcceptance' ? 'PendingAcceptance'
                       : user.externalUserState === 'Accepted' ? 'Accepted' : 'Unknown',
    };
  }

  /** Filter helpers for the UI. */
  function passesFilter(g, f) {
    // f: { search, minInactive, redemption, enabled, neverSignedIn, maxAge, minAge }
    if (f.search) {
      const q = f.search.toLowerCase();
      const hay = `${g.displayName || ''}\n${g.userPrincipalName || ''}\n${g.mail || ''}\n${g.companyName || ''}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    if (f.redemption && f.redemption !== 'any' && g.redemptionState !== f.redemption) return false;
    if (f.enabled === 'enabled' && !g.accountEnabled) return false;
    if (f.enabled === 'disabled' && g.accountEnabled) return false;
    if (f.neverSignedIn === 'never' && !g.neverSignedIn) return false;
    if (f.neverSignedIn === 'signed' && g.neverSignedIn) return false;
    if (typeof f.minInactive === 'number' && (g.inactiveDays == null || g.inactiveDays < f.minInactive)) return false;
    if (typeof f.minAge === 'number' && (g.accountAgeDays == null || g.accountAgeDays < f.minAge)) return false;
    if (typeof f.maxAge === 'number' && (g.accountAgeDays == null || g.accountAgeDays > f.maxAge)) return false;
    return true;
  }

  function fmtDate(d) {
    if (!d) return '';
    const dt = d instanceof Date ? d : new Date(d);
    if (isNaN(dt.getTime())) return '';
    return dt.toISOString().substring(0, 10);
  }

  window.Guests = { lastSignInDate, daysBetween, enrich, passesFilter, fmtDate };
})();
