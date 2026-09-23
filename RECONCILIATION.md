# Group Report Factory reconciliation gate

Status: OPEN — reconciliation is the first and only task on this branch.

## First task

Resolve the differences between the existing GRF design branch and the incoming GRF implementation packet before doing any other work.

- Destination baseline: the existing GRF design branch at the branch point for this reconciliation.
- Incoming material: the GRF-specific source, tests, build changes and safe assets from the platform-side work packet.
- Do not copy governance notes, project records, human-document work, client exports, participant data, screenshots or unreviewed reference material into this public repository.
- Preserve the source-to-deploy contract: source changes must be built into the deploy surface and independently checked.

## Branch gate

Until the reconciliation and privacy review are resolved:

- no feature development, refactoring, deployment or merge to `main`;
- no publication of unreviewed assets or generated Office material;
- no client, participant, project or real-person identity may enter this repository;
- no claim that the GRF output is accepted or production-ready;
- every conflict disposition must identify the winning behavior, the preserved behavior, and the verification that proves the result.

The fresh Engineer session that resolves this branch must update this file with the conflict disposition, privacy checks and source-to-deploy verification. A separate Inspector session must review the resolved diff before any merge or branch cleanup.
