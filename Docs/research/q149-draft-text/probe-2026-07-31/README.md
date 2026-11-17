# The D49 probe of 2026-07-31, recovered

The probe the D49 design rests on - "the invisible compose surface" and the lifetime pin - run on
2026-07-31 on the maintainer's workstation: Office LTSC 2021, 16.0.14334, add-ins loaded, every run
a separate process. Its folder in Downloads was cleared some time after; on 2026-10-11 its source and
every result were recovered from that session's transcript, and they are committed here because that
transcript was the only copy left.

**Identifiers are replaced**: the test mailbox the probe wrote to is `hub@example.invalid`, the two
other accounts it read are `account-a@example.invalid` and `account-b@example.invalid`, the external
address the wire check sent to is `external@example.invalid`, and the user's profile folder is
`<user>`. Nothing else was changed.

**Archived with a `.txt` suffix**: these are records, not tools. Under their own names the
repository's PowerShell 5.1 guard and the CodeQL scan would treat them as live code, and they must not
run again - they wrote to and sent from real mailboxes.

| Path | What it is |
| --- | --- |
| `source/Program.cs.txt` | The probe: one window state per run (`headless`, `explorer-nodisplay`, `inspector-activate`, `pinned-hidden-activate`, `htmlnative` ...), a tagged draft, the editor read, the close, a 300 ms liveness check. Late-bound COM on a pumped STA. |
| `source/Sta.cs.txt`, `source/Win.cs.txt` | Its STA runner (with the retry message filter) and its window enumeration. |
| `source/struct-Program.cs.txt` | The structure probe that read the HTML each state produced. |
| `source/d49-wire-verify.ps1.txt`, `source/d49-verify-received.ps1.txt` | The round trip of a D49 draft through the shipped server, to the wire and back. |
| `source/quit-outlook.ps1.txt` | How the runs ended Outlook between them. |
| `source/edits.txt` | Every edit made to the probe between runs, in order. |
| `results/all-results.txt` | The 27 raw results, as the probe printed them. |
| `results/runs.tsv` | One row per run (26), with the fields the results carry. |
| `results/trials-table.txt` | Every trial classified for the lifetime question by the 2026-10-11 research: cold or attached, the close, alive at 300 ms. |

How many runs stand behind each row of the D49 matrix, and where the record overstates, is in
[the Q149 README](../README.md#corrections-to-the-d49-record): 22 valid runs in all, and three of
them were not the fresh cold starts the record says every run was.
