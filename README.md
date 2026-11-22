# spfx-listview-actions

## Summary

A SharePoint Framework (SPFx) **ListView Command Set** (`BtnFlowTriggerCommandSet`) that adds command bar actions to modern SharePoint lists and document libraries to trigger a **Power Automate** HTTP webhook with selected item metadata and/or queue a **Power BI** semantic model refresh via `PowerBiRefreshService`.

Built with TypeScript, SPFx (`@microsoft/sp-listview-extensibility`, `@microsoft/sp-http`, `@microsoft/sp-dialog`), Rush Stack **Heft**, and **Jest** for local unit testing.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.24.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- **Node.js** (LTS version supported by your SPFx toolchain) and **npm**
- **Power Automate:** An HTTP-triggered flow URL if using `actionMode` set to `"flow"` or `"both"`
- **Power BI (optional):** Target workspace ID and dataset/semantic model ID, plus tenant/admin API permission approval (`Dataset.ReadWrite.All` or equivalent Power BI Service scope) when using `AadHttpClient` in SharePoint

## Solution

| Solution | Author(s) |
| --- | --- |
| `src/extensions/btnFlowTrigger` | Ross Bellingham |

## Version history

| Version | Date | Comments |
| --- | --- | --- |
| 1.0 | November 2026 | Initial release with Power Automate webhook trigger, PowerBiRefreshService, and local Jest unit tests |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure you are in the project root folder
- Run the following commands:
  - `npm install`
  - `npm test` (runs local Jest unit tests without requiring a SharePoint tenant)
  - `heft start` or your configured Heft/build script (see `package.json`)

Other build commands can be listed using `npx heft --help`.

## Features

- **Dual Action Modes (`actionMode`):** Configure the command set for `"flow"`, `"powerbi"`, or `"both"` actions.
- **Selection Control (`requireSelection`):** Require at least one selected list/library row for the flow trigger action (enabled by default), while allowing dataset refreshes as configured.
- **Power Automate Webhook Payload:** Maps selected SharePoint rows (falling back to `FileLeafRef` when `Title` is empty) and sends `triggeredAtUtc`, `triggeredBy`, `sharePointContext`, and `selectedItems` as JSON via `HttpClient`.
- **Power BI Dataset Refresh Service (`PowerBiRefreshService`):** Encapsulates configuration validation, scope normalization (stripping trailing `/.default`), `AadHttpClient` initialization, POST refresh requests, and `202 Accepted` handling.
- **Local Unit Testing:** Unit tests in `PowerBiRefreshService.spec.ts` run locally in Node with Jest using mocked SPFx HTTP/extension dependencies.

### Extension Properties

| Property | Type | Default | Description |
| --- | --- | --- | --- |
| `actionMode` | `'flow' \| 'powerbi' \| 'both'` | `'flow'` | Controls which command bar button(s) are visible |
| `flowWebhookUrl` | `string` | `undefined` | HTTP POST trigger URL for Power Automate |
| `powerBiWorkspaceId` | `string` | `undefined` | Target Power BI workspace (group) ID |
| `powerBiDatasetId` | `string` | `undefined` | Target Power BI dataset ID |
| `powerBiApiScope` | `string` | `'https://analysis.windows.net/powerbi/api/.default'` | Power BI API scope/resource URI |
| `requireSelection` | `boolean` | `true` | Hides the flow command when no rows are selected if `true` |

## Roadmap / Planned Extensions

As `spfx-listview-actions` grows beyond `BtnFlowTriggerCommandSet` and `PowerBiRefreshService`, planned additions follow the same pattern of lightweight command sets backed by testable service modules:

- **Additional ListView Command Sets:** Dedicated command sets for common list/library operations (such as bulk status updates, metadata actions, or document lifecycle triggers) without bloating `BtnFlowTriggerCommandSet`.
- **Reusable Microsoft 365 & Power Platform Services:** Standalone service classes wrapping Graph, SharePoint REST, or Power Platform APIs with explicit inbound/outbound boundary handling and local Jest unit tests.
- **CI/CD Automation:** GitHub Actions workflow running `npm test` and `npm run build` on pull requests and pushes to `main`.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [SharePoint Framework ListView Command Sets](https://learn.microsoft.com/sharepoint/dev/spfx/extensions/get-started/building-simple-cmdset-from-scratch)
- [Connect to Azure AD-secured APIs with AadHttpClient](https://learn.microsoft.com/sharepoint/dev/spfx/use-aadhttpclient)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)