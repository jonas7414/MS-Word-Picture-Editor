# GitHub Release Automation

## Behavior

On every push to `master`, the `Release` workflow:

1. Installs dependencies.
2. Runs tests.
3. Builds the GitHub Pages package.
4. Reads `version` from `package.json`.
5. Creates a GitHub release named `v<version>` if it does not already exist.

Release assets:

```text
word-pic-editor-v<version>.zip
manifest.xml
```

If the release tag already exists, the workflow skips release creation.

## Publish a New Release

1. Update `package.json`:

   ```json
   {
     "version": "0.1.1"
   }
   ```

2. Commit and push to `master`.

3. Wait for the `Release` workflow to finish.

4. Download assets from GitHub:

   ```text
   https://github.com/jonas7414/MS-Word-Picture-Editor/releases
   ```

Use the release `manifest.xml` for Word sideloading, or keep using the GitHub Pages manifest:

```text
https://jonas7414.github.io/MS-Word-Picture-Editor/manifest.xml
```
