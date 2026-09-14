# P4.3 Build Fix

The P4/P4.1 build failed under `noUncheckedIndexedAccess` because array index access is typed as possibly undefined.

P4.3 fixes:
- register column widths
- register page access
- party Up / Down swap access

A source-only strict TypeScript check passes locally. Test files were excluded from that local check only because the local container does not have the Vitest package installed; Vercel installs project dependencies normally.
