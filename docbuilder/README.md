# Reports UI + DocumentBuilder

This folder is reserved for a local DocumentBuilder runtime used by `reports-ui`.

## Target layout

- `reports-ui/docbuilder/bin/docbuilder.exe`
- `reports-ui/docbuilder/bin/*.dll` and runtime files
- `reports-ui/docbuilder/scripts/*.docbuilder`

## Native API bridge

After rebuilding DesktopEditors with the `docbuilder:run` bridge, `reports-ui` exposes:

- `window.ReportsDocBuilder.probe(options?)`
- `window.ReportsDocBuilder.run(payload)`

Both methods return a Promise.

Example:

```js
await window.ReportsDocBuilder.probe();

await window.ReportsDocBuilder.run({
  script: 'reports-ui/docbuilder/scripts/sample_table.docbuilder',
  argument: {
    output: 'C:/Temp/report.xlsx',
    title: 'Generated from reports-ui'
  },
  timeoutMs: 120000
});
```

## Runtime sync script

Use `reports-ui/tools/sync-docbuilder-runtime.ps1` to copy a built DocumentBuilder runtime
into `reports-ui/docbuilder/bin`.
