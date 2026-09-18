# gas-spread-chat

## Deployment

The GAS deployment target is defined only by the tracked `.clasp.json` file. The same file is used by local clasp pushes and by GitHub Actions, so the Script ID does not have a second copy in Actions inputs, secrets, or environment variables.

When changing the development spreadsheet, update `.clasp.json` locally and commit that change before pushing GAS code. For local deployment, use:

```bash
bash scripts/clasp-push-safe.sh
```

The wrapper refuses to push when `.clasp.json` is missing, untracked, staged, or modified relative to `HEAD`. Direct `clasp push` still bypasses that guard.

Pushing or merging to `main` triggers the GitHub Actions deployment workflow. The workflow reads the committed `.clasp.json` directly. Until a tracked `.clasp.json` exists, the workflow exits successfully without deploying.

The workflow requires a repository secret named `CLASPRC_JSON`. Set it to the full JSON contents of your local `~/.clasprc.json`. A base64-encoded copy of the same JSON is also accepted.
