Rota app safe patch - site-specific Front Desk bands

This package keeps the last-working app import name:
  rota_engine_v37_19_TARGETS_FINAL.py

Latest fixes included:
- SLGP Front Desk locked to: 08:00-11:00, 11:00-13:00, 13:00-16:00, 16:00-18:30.
- JEN and BGS Front Desk locked to: 08:00-10:30, 10:30-13:00, 13:00-16:00, 16:00-18:30.
- Front Desk is checked/locked as a first-priority rule and breaks must move around it.
- No duplicate 30-minute breaks for the same person/day.
- Email daily priority remains in place, with break allowed within email cover.
- BGS triage admin remains included as a task.
- Misc remains last resort, with EMIS/Docman target filling prioritised.
- Heatmaps/target checks remain included in the workbook output.
