# Design QA — role-first workspace 5.0.0

final result: passed

## Target and evidence

Selected direction: third image, “Рабочие разделы”. Source board:
`exec-8aa4c48f-e842-41af-8406-81bf0bfe62d0.png` (1672 × 941).
Source desktop crop: 23,48 / 1265 × 868; mobile crop: 1303,48 / 347 × 869.
The generated board has approximate rather than exact device proportions. Crops
were normalized to actual CSS viewports, without treating this distortion as a bug.

Rendered implementation: localhost existing production app with synthetic fixture
data only, kitchen selected, yesterday, owner signed in. No production grants changed.

- Desktop: 1440 × 1024 CSS px, 1× capture, `/tmp/personnel-desktop-final.png`.
- Mobile: 390 × 844 CSS px, 1× capture, `/tmp/personnel-mobile-final.png`.
- Combined source/implementation inputs opened and inspected:
  `/tmp/personnel-compare-desktop.png`, `/tmp/personnel-compare-mobile.png`.
- Earlier captures: `/tmp/personnel-design-desktop-v2.png`,
  `/tmp/personnel-design-mobile.png`; users: `/tmp/personnel-users.png`.

## Iteration and fixes

Initial result blocked: desktop profile/version overlap; mobile date input overflow;
wide legacy table; excessive control height; mobile numeric alignment; hidden-text
icon controls without names. Fixed by moving version to footer, constrained date
popover, role-specific columns, condensed spacing, explicit right alignment and
accessible filter/more names. Re-captured both viewports and compared together with
the source after these changes. No remaining P0/P1/P2 visual issues identified.

## Fidelity surfaces

- Typography: system sans, clear title/stat/body hierarchy. Desktop deliberately
  denser than the nine-row illustrative source to show all 18 restaurants plus
  production. Mobile names 13px, numeric values 14px, full-width tappable name rows.
- Rhythm/layout: slim sidebar and compact controls, role underline, two summary
  metrics, four information columns. At 1440 × 1024 all 19 rows are visible;
  shorter screens scroll vertically. Mobile has bottom navigation and no wide table.
- Colors: white background, muted neutral dividers and secondary labels, green
  active states, amber only for actual warnings. No decorative gradients added.
- Assets: existing application team icon retained intentionally instead of the
  generated lettermark. Feather 4.29.2 provides consistent UI icons, MIT included.
- Copy/content: illustrative numbers are replaced by actual report fields in
  production, not hardcoded mock data. Production-only factory displays staffing;
  no fabricated revenue or zero substitution. Longer canonical site names retained.

Focused checks: mobile text/numeric columns and desktop table are readable in the
combined captures at 1×; role/date/employee controls checked separately in the UI.
Search is an intentional addition; account breadcrumb is retained for orientation.

## Functional/accessibility checks

Browser: role switching, enterprise dialog, highlighted reason on employee,
Google users auto-loading/grouping, restricted employee showing only one allowed
site and kitchen, mobile native controls. No browser console errors observed.
Native dialog provides focus containment/Escape; keyboard focus outlines retained
on controls; headings no longer receive a large programmatic-focus rectangle.

Automated regression suite passed: attendance schedule/resume and throttling,
PWA installed-state detection, report authorization/category revenue, production
exclusion, weighted ratios and missing days, access autoload deduplication and
late-session response guard, Saby auth, workspace routing, Moscow-yesterday default.

Remaining P3: mobile list deliberately shows fewer rows than the illustrative
image because full canonical names and real warning reasons must remain readable.
Very long enterprise/employee lists scroll; “one screen” is not forced by shrinking
all content. Data loading/server synchronization behavior was not changed here.
