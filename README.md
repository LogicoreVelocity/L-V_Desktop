# Logicore Excel Manager

Logicore Excel Manager is a local-first desktop app for replacing scattered HR, finance, asset, payroll, and document spreadsheets with one structured Excel import/export tool.

It is intended for small employers, fleet operators, manpower teams, and delivery companies that need payroll preparation without sending employee data to a hosted service.

## Included Modules

- Employees / drivers
- Finance register
- Asset register
- Document / PDF metadata register
- Payroll inputs
- Advances and deductions
- Payroll export
- WPS/SIF export validation
- Salary slip PDFs
- Local JSON backup and restore

## Local Data Position

This app stores working data in the user's browser/Electron local storage and in files the user explicitly exports. Logicore Velocity does not receive, host, own, or process tenant employee/payroll data in this local-first release.

There is no API server, hosted login, PostgreSQL database, Docker stack, or cloud tenant layer in this repository. Those layers can be added later as a separate backend integration without changing the public local-first operating model.

## Import And Export

The app uses the `.xlsx` file format for Excel import/export. PDF files are registered as local document metadata; the app does not store full PDF contents in localStorage.

## License And Branding Notes

Electron applications include `LICENSE.electron.txt` and `LICENSES.chromium.html` in the packaged output. Those files are expected. They are third-party notices for Electron, Chromium, and bundled open-source components. Keep them in the installer/unpacked app.

Using the app name `Logicore Excel Manager` is separate from Electron/Chromium licensing, provided the app does not imply Electron or Chromium endorsement and preserves their notices.

## Development

Install dependencies:

```powershell
npm.cmd install
```

Run the local web app:

```powershell
npm.cmd run dev
```

Build and validate:

```powershell
npm.cmd test
npm.cmd run build
npm.cmd run lint
npm.cmd audit --omit=dev
```

Run the Electron app locally:

```powershell
npm.cmd run electron:dev
```

Build the Windows installer:

```powershell
npm.cmd run electron:build
```

The Windows installer is written to:

```text
C:\Users\sahal\Desktop\Project\logicore-excel-manager-artifacts
```

## Readiness Checks

The automated test suite covers payroll calculation, WPS/SIF validation, Excel file preflight checks, worksheet parsing, required import columns, duplicate employee keys, duplicate payroll inputs, and deduction import validation.

Run all checks before publishing a release:

```powershell
npm.cmd test
npm.cmd run lint
npm.cmd run build
npm.cmd audit --omit=dev
```
