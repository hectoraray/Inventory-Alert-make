# Test plan: Make scenario

The formulas are verified automatically (`scripts/verify_against_n8n.py`). This checklist exercises the **Make scenario** itself: the connections, the router, the aggregator and the write-back. Run it once with a fresh copy of the sample sheet and tick each box only when the result matches.

Start state: sheet imported from `data/inventory_make_template.xlsx`, column `last_alert_status` as in the file (SKU-003, SKU-010 = `LOW`; SKU-006 = `CRITICAL`; SKU-021 = `OK`).

| # | Action | Expected result | Done |
|---|---|---|---|
| 1 | Run the scenario once | **One** Telegram message listing 7 items, most severe first: 1 OUT (SKU-006), 4 CRITICAL (SKU-004, 008, 014, 022), 2 data issues (SKU-017, SKU-019). Column G now holds the status of exactly those 7 rows | [ ] |
| 2 | Run again immediately | No Telegram message. No cell changes. The run stops after the search module | [ ] |
| 3 | **Escalation:** set SKU-003 stock to `5` (was `LOW`) and run | One message with SKU-003 as CRITICAL; its column G becomes `CRITICAL` | [ ] |
| 4 | **Silent recovery:** set SKU-004 stock to `200` and run | No message; SKU-004 column G becomes `OK`. No error from an empty message (this exercises the "text is not empty" filter) | [ ] |
| 5 | **Alert after recovery:** set SKU-004 stock back to `3` and run | A message with SKU-004 as CRITICAL (the state was not stuck on the old value) | [ ] |
| 6 | **Data fix:** type a number in SKU-017 stock and run | No message; its column G is rewritten to its new status | [ ] |
| 7 | **Failure path:** put a wrong Telegram token in the connection, make a row escalate (e.g. SKU-005 stock `1`), run | The scenario reports an error. SKU-005 column G is **unchanged** (no state was lost) | [ ] |
| 8 | Restore the token and run | The SKU-005 alert is delivered; column G updates | [ ] |
| 9 | **Row limit:** (optional) lower the search limit to `2`, reset column G, run twice | First run alerts 2 rows (most severe first); second run picks up the next ones | [ ] |

Test 7 and 8 check the key claim of the design: if sending fails, the alert is not lost. If SKU-005 column G *does* change in test 7, the router is running the update route before the notification route, or the Break error handler is missing on the Telegram module. Fix that (notification route first, Break handler added) and re-test.

Add a screenshot of the scenario and of the Telegram message to `docs/screenshots/` once all boxes are ticked.
