# gas-spread-chat

## Deployment

The GitHub Actions deployment workflow requires a repository secret named `CLASPRC_JSON`.
Set it to the full JSON contents of your local `~/.clasprc.json` before running the manual `Deploy to GAS` workflow.
The workflow now also accepts a base64-encoded copy of that same JSON.

If the workflow says `secrets.CLASPRC_JSON` is empty, GitHub Actions did not receive a usable value for that exact secret name.
That is different from malformed JSON or a character-encoding problem, which will fail later with a JSON parsing error instead.
