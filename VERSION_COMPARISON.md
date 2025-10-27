# Version Comparison for Session Insight Rail Updates

The four candidate revisions differ mainly in how fully they integrate the new
right-hand rail into the DOW 30 Tracker experience and supporting tooling.

## Recommendation
Version 4 is the strongest overall. It combines the richer rail telemetry from
Versions 2 and 3 with the pipeline and build improvements introduced in Version
1 while also ensuring the new UI is synchronized with capture hooks, feature
defaults, and export updates. The addition of reusable release metadata plus the
streamlined build instructions makes it the most complete package for day-to-day
use.

## Rationale by Version
- **Version 1** – Introduces the InsightPanel concept but leaves the toolbar and
  capture plumbing largely untouched, so the rail risks drifting out of sync
  with live updates.
- **Version 2** – Renames the component to the Roadmap Rail and improves
  backfill/export coordination, but it omits the broader release-tooling
  refresh.
- **Version 3** – Adds countdowns and breadth tracking while wiring the rail
  deeper into the capture lifecycle, yet it still lacks the reusable metadata
  for the build script and installer.
- **Version 4** – Builds on the earlier iterations to deliver the most polished
  UI integration and completes the tooling updates, providing a better balance
  between user experience and maintainability.

