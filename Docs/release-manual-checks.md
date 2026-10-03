# Manual checks before a release

**What this is.** A short list of checks the maintainer runs BY HAND before a release, for the
behaviour no automated test can reach: writes that need an Exchange profile. Decided 2026-10-03 (Q74,
"D1 + D2"). The test machines have no Exchange (AGENTS.md, Dependencies), and the maintainer's
workstation is read-only for live tests (Q72), so send-on-behalf, delegate-mailbox drafts and the
like are covered by unit checks of their decisions (`T1/ExchangeWritePathTests`, D1) and by this list
(D2), and by nothing else.

**Who runs it: the maintainer, himself.** Every step below writes to his real mailbox. An agent never
performs any of it, never on his behalf, and never "to save time" - mailbox-safety rules 1 to 3 apply in
full. An agent may ask whether it has been done; that is all.

**The installer is a separate check, and never on the workstation** (Q97 2(b), 2026-10-03). The release
candidate's installer is tested on a test VM: `Testbed/host/Publish-AddInPayload.ps1` builds it from the
commit and `Testbed/guest/Install-OutlookAIAddIn.ps1` installs it on a guest. The workstation keeps the
copy-and-register route, `Tools/Switch-AddInBuild.ps1`, for running a dev build there when he asks for one.

## Before you start

- **The server answering must be the build under test.** Call `outlook_health` and read `runningFrom`
  and the version, and write them down. There is no route yet for putting a release candidate's MCP
  server in front of a session on the workstation without installing it (`Tools/Switch-AddInBuild.ps1`
  moves the add-in only, and never the server - `TODO.md`). Until there is, run this list right after
  installing the release, and keep the previous installer: a failure here is a reason to put it back.
- **What you need:** a shared mailbox S on the profile (auto-mapped, not an account of its own) on which
  you have Full Access and Send on Behalf; and an address X on which you do NOT have Send on Behalf.
- **Tag every subject** `[OutlookAI-D2 <version>]` and send only to yourself.

## The checks

Each step: what to do, what you should see, and what counts as a failure. Write down what the tool
answered whenever a step fails - the answer is the evidence.

1. **Plain send from your own mailbox.** `new_draft` to yourself, then `send`, then `send` again with the
   token. *See:* step 1 names you as `account`; step 2 says `sent` with `accountVerified: true`; the mail
   arrives from you. *Fail:* `no_sending_account` or `identity_verification_failed`, or it arrives from any
   other account.
2. **Send on behalf, permitted.** The same, with `sent_on_behalf_of` = S on both calls. *See:* it arrives
   as you on behalf of S, and the result's `sentOnBehalfOf` is S. *Fail:* no on-behalf line, sent AS S, or
   the result disagrees with what arrived.
3. **Send on behalf, not permitted.** The same, with X. *See:* a non-delivery report saying you cannot send
   on behalf of X, and nothing delivered. *Fail:* the mail is delivered; or nothing comes back at all while
   the tool said `sent`. Record exactly what the tool said.
4. **Reply to a mail in S**, with display on. *See:* `store` is S and `accountResolved: false`; the draft is in
   S's Drafts, threaded, with the quote and the signature intact. *Fail:* the draft is not where the result
   says, `accountResolved: true`, or it is unthreaded or the body is lost.
5. **`send` that draft**, without and then with `sent_on_behalf_of` = S. *See:* both refuse with
   `no_sending_account` and issue no token. *Fail:* any token issued, or anything leaves.
6. **`update_draft`, then `discard_draft`, then undo with `move_mail`.** *See:* the revision intact; note which
   Deleted Items holds the discarded draft, and whether that matches what the result says. *Fail:* it is in
   neither Deleted Items, or an undo the result offered does not work.
7. **`archive_mail` the test mail in S, then undo.** *See:* it lands in S's own Archive and `createdFolders`
   is absent. *Fail:* it lands in your own mailbox, a new folder appears, or the undo fails.
8. **Clean up.** Delete the tagged items yourself, in Outlook. *See:* the Outbox empty, nothing tagged left in
   S, and one audit-log line for every write step above. *Fail:* anything stuck in the Outbox, or a write
   with no audit line.

A failure here is a release blocker until it is understood: these are exactly the paths no automated test
in this repository can take.
