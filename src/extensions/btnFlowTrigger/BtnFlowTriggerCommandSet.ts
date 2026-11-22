import { Log } from '@microsoft/sp-core-library';
import { Dialog } from '@microsoft/sp-dialog';
import {
  HttpClient,
  HttpClientResponse,
  IHttpClientOptions
} from '@microsoft/sp-http';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import * as strings from 'BtnFlowTriggerCommandSetStrings';
import {
  PowerBiRefreshService,
  PowerBiNotifyOption
} from './services/PowerBiRefreshService';
import { SharePointBatchUpdateService } from './services/SharePointBatchUpdateService';
import { GraphActionService } from './services/GraphActionService';
import { DocumentLifecycleService } from './services/DocumentLifecycleService';

export interface IBtnFlowTriggerCommandSetProperties {
  actionMode?: 'flow' | 'powerbi' | 'both' | 'all';
  webhookUrl?: string;
  workspaceId?: string;
  datasetId?: string;
  notifyOption?: PowerBiNotifyOption;
  batchUpdateFieldName?: string;
  batchUpdateFieldValue?: string;
  batchSize?: number;
  teamId?: string;
  channelId?: string;
  archiveFolderServerRelativeUrl?: string;
  requireSelection?: boolean;
}

const LOG_SOURCE: string = 'BtnFlowTriggerCommandSet';

// Command IDs must match the command definitions in BtnFlowTriggerCommandSet.manifest.json.
const CMD_TRIGGER_FLOW = 'CMD_TRIGGER_FLOW';
const CMD_REFRESH_POWERBI = 'CMD_REFRESH_POWERBI';
const CMD_BATCH_UPDATE = 'CMD_BATCH_UPDATE';
const CMD_POST_TEAMS_MESSAGE = 'CMD_POST_TEAMS_MESSAGE';
const CMD_ARCHIVE_DOCUMENTS = 'CMD_ARCHIVE_DOCUMENTS';

export interface ISelectedItemPayload {
  id: string;
  title: string;
  fileLeafRef?: string;
  contentType?: string;
}

export interface IFlowTriggerPayload {
  triggeredAtUtc: string;
  triggeredBy: {
    displayName: string;
    email: string;
    loginName: string;
  };
  sharePointContext: {
    webAbsoluteUrl: string;
    listId?: string;
    listTitle?: string;
  };
  selectedItems: ISelectedItemPayload[];
}

/**
 * Minimal row shape so row mapping can be unit-tested without SharePoint's internal RowAccessor.
 */
export interface IListRowReader {
  getValueByName(name: string): unknown;
}

/**
 * Pure helper to map SharePoint list rows into normalized item objects.
 * Inbound: raw SharePoint list row objects.
 * Outbound: clean array of ISelectedItemPayload objects for serialization.
 */
export function mapSelectedRows(rows: readonly IListRowReader[]): ISelectedItemPayload[] {
  return rows.map((row) => ({
    id: String(row.getValueByName('ID') ?? ''),
    title: String(row.getValueByName('Title') ?? row.getValueByName('FileLeafRef') ?? ''),
    fileLeafRef: row.getValueByName('FileLeafRef') ? String(row.getValueByName('FileLeafRef')) : undefined,
    contentType: row.getValueByName('ContentType') ? String(row.getValueByName('ContentType')) : undefined
  }));
}

/**
 * Pure helper to build the JSON body sent to Power Automate.
 * Inbound: timestamp, user/web/list context metadata, and mapped selected items.
 * Outbound: structured IFlowTriggerPayload object.
 */
export function buildFlowTriggerPayload(params: {
  triggeredAtUtc: string;
  user: { displayName?: string; email?: string; loginName?: string };
  webAbsoluteUrl: string;
  listId?: string;
  listTitle?: string;
  selectedItems: ISelectedItemPayload[];
}): IFlowTriggerPayload {
  return {
    triggeredAtUtc: params.triggeredAtUtc,
    triggeredBy: {
      displayName: params.user.displayName ?? '',
      email: params.user.email ?? '',
      loginName: params.user.loginName ?? ''
    },
    sharePointContext: {
      webAbsoluteUrl: params.webAbsoluteUrl,
      listId: params.listId,
      listTitle: params.listTitle
    },
    selectedItems: params.selectedItems
  };
}

export default class BtnFlowTriggerCommandSet
  extends BaseListViewCommandSet<IBtnFlowTriggerCommandSetProperties> {

  private _flowCommand?: Command;
  private _pbiCommand?: Command;
  private _batchUpdateCommand?: Command;
  private _teamsMessageCommand?: Command;
  private _archiveDocumentsCommand?: Command;
  private _powerBiService?: PowerBiRefreshService;

  // Runs once when the ListView Command Set is initialized ("onload" setup).
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initializing BtnFlowTriggerCommandSet');

    this._flowCommand = this.tryGetCommand(CMD_TRIGGER_FLOW);
    this._pbiCommand = this.tryGetCommand(CMD_REFRESH_POWERBI);
    this._batchUpdateCommand = this.tryGetCommand(CMD_BATCH_UPDATE);
    this._teamsMessageCommand = this.tryGetCommand(CMD_POST_TEAMS_MESSAGE);
    this._archiveDocumentsCommand = this.tryGetCommand(CMD_ARCHIVE_DOCUMENTS);

    // INBOUND: Extension properties from manifest, serve.json, or SharePoint deployment config.
    const workspaceId = this.properties.workspaceId?.trim();
    const datasetId = this.properties.datasetId?.trim();

    if (workspaceId && datasetId) {
      // OUTBOUND: Passing Power BI target config + AadHttpClientFactory into PowerBiRefreshService.
      this._powerBiService = new PowerBiRefreshService(
        {
          workspaceId,
          datasetId,
          notifyOption: this.properties.notifyOption
        },
        this.context.aadHttpClientFactory
      );

      try {
        // OUTBOUND: Requests AAD client initialization for Power BI API scope.
        await this._powerBiService.initialize();
      } catch (err) {
        // INBOUND: Initialization error caught so command set still loads; retry happens on execute.
        Log.warn(LOG_SOURCE, `PowerBiRefreshService init deferred: ${String(err)}`);
      }
    } else {
      Log.warn(LOG_SOURCE, 'Power BI workspaceId or datasetId not configured; skipping PBI init.');
    }

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    // INBOUND: Initial selected row count from SharePoint list view state.
    this._updateCommandVisibility(this.context.listView.selectedRows?.length ?? 0);
  }

  public onDispose(): void {
    this.context.listView.listViewStateChangedEvent.remove(this, this._onListViewStateChanged);
  }

  // Runs whenever list selection or view state changes.
  private _onListViewStateChanged = (): void => {
    // INBOUND: Updated selected row count from list view event.
    const selectedCount = this.context.listView.selectedRows?.length ?? 0;
    // OUTBOUND: Updates command bar button visibility flags on the Command objects.
    this._updateCommandVisibility(selectedCount);
  };

  // Runs when a command bar button is clicked ("onsubmit" / execution handler).
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    // INBOUND: Selected rows passed by SharePoint command bar execution event.
    const rawRows = event.selectedRows || [];
    const selectedItems = mapSelectedRows(rawRows);

    try {
      switch (event.itemId) {
        case CMD_TRIGGER_FLOW: {
          if (!this.properties.webhookUrl) {
            await Dialog.alert(strings.MissingConfigMessage);
            return;
          }
          // OUTBOUND: Sends selectedItems + SharePoint context to Power Automate HTTP webhook.
          await this._triggerPowerAutomateFlow(this.properties.webhookUrl, selectedItems);
          await Dialog.alert(strings.FlowSuccessMessage.replace('{0}', String(selectedItems.length)));
          break;
        }

        case CMD_REFRESH_POWERBI: {
          if (!this._powerBiService || !this._powerBiService.isConfigured()) {
            await Dialog.alert(strings.MissingConfigMessage);
            return;
          }
          // OUTBOUND: Invokes Power BI dataset refresh; INBOUND: receives HTTP status code (e.g. 202).
          const statusCode = await this._powerBiService.triggerRefresh();
          await Dialog.alert(strings.PbiRefreshAcceptedMessage.replace('{0}', String(statusCode)));
          break;
        }

        case CMD_BATCH_UPDATE: {
          await this._executeBatchUpdate(selectedItems);
          break;
        }

        case CMD_POST_TEAMS_MESSAGE: {
          await this._executePostTeamsMessage(selectedItems);
          break;
        }

        case CMD_ARCHIVE_DOCUMENTS: {
          await this._executeArchiveDocuments(selectedItems);
          break;
        }

        default:
          throw new Error(`Unknown command ID: ${event.itemId}`);
      }
    } catch (error: unknown) {
      // INBOUND: Error thrown by service/HTTP call or validation.
      // OUTBOUND: Logs error and surfaces user-facing message via SharePoint Dialog.
      const message = error instanceof Error ? error.message : String(error);
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(message));
      await Dialog.alert(strings.ActionFailedMessage.replace('{0}', message));
    }
  }

  private async _executeBatchUpdate(selectedItems: ISelectedItemPayload[]): Promise<void> {
    const fieldName = this.properties.batchUpdateFieldName?.trim();
    const fieldValue = this.properties.batchUpdateFieldValue;
    const listTitle = this.context.pageContext.list?.title;

    if (!fieldName || fieldValue === undefined || !listTitle) {
      await Dialog.alert(strings.MissingConfigMessage);
      return;
    }

    const itemIds = this._parseNumericIds(selectedItems);
    if (itemIds.length === 0) {
      await Dialog.alert(strings.MissingConfigMessage);
      return;
    }

    // OUTBOUND: Instantiates SharePointBatchUpdateService with site/list context and chunk size.
    const batchService = new SharePointBatchUpdateService(this.context.spHttpClient, {
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listTitle,
      batchSize: this.properties.batchSize
    });

    // OUTBOUND: Updates field across selected items; INBOUND: returns per-item summary.
    const summary = await batchService.updateItems(itemIds, { [fieldName]: fieldValue });

    if (summary.failed > 0) {
      const failures = summary.results
        .filter((r) => !r.success)
        .map((r) => `ID ${r.itemId}: ${r.error}`)
        .join('; ');
      throw new Error(`Batch update completed with ${summary.failed} failure(s): ${failures}`);
    }

    await Dialog.alert(`Updated ${summary.succeeded} of ${summary.total} item(s).`);
  }

  private async _executePostTeamsMessage(selectedItems: ISelectedItemPayload[]): Promise<void> {
    const teamId = this.properties.teamId?.trim();
    const channelId = this.properties.channelId?.trim();

    if (!teamId || !channelId) {
      await Dialog.alert(strings.MissingConfigMessage);
      return;
    }

    // OUTBOUND: Acquires Graph-scoped AadHttpClient via SPFx factory.
    // INBOUND: Authenticated AadHttpClient for Microsoft Graph.
    const aadClient = await this.context.aadHttpClientFactory.getClient('https://graph.microsoft.com');
    const graphService = new GraphActionService(aadClient);

    const listTitle = this.context.pageContext.list?.title ?? 'SharePoint List';
    const itemSummary = selectedItems
      .map((item) => item.title || item.fileLeafRef || `ID ${item.id}`)
      .join(', ');
    const bodyHtml = `<p>Action triggered on <strong>${listTitle}</strong> (${selectedItems.length} item(s)): ${itemSummary}</p>`;

    // OUTBOUND: Posts HTML message to Teams channel; INBOUND: created Graph message ID.
    const messageId = await graphService.postChannelMessage({
      teamId,
      channelId,
      bodyHtml
    });

    Log.info(LOG_SOURCE, `Posted Teams message ID: ${messageId}`);
    await Dialog.alert(`Posted update to Microsoft Teams (${selectedItems.length} item(s)).`);
  }

  private async _executeArchiveDocuments(selectedItems: ISelectedItemPayload[]): Promise<void> {
    const archiveFolder = this.properties.archiveFolderServerRelativeUrl?.trim();
    const libraryTitle = this.context.pageContext.list?.title;

    if (!archiveFolder || !libraryTitle) {
      await Dialog.alert(strings.MissingConfigMessage);
      return;
    }

    const itemIds = this._parseNumericIds(selectedItems);
    if (itemIds.length === 0) {
      await Dialog.alert(strings.MissingConfigMessage);
      return;
    }

    // OUTBOUND: Instantiates DocumentLifecycleService for library archive moves.
    const lifecycleService = new DocumentLifecycleService(this.context.spHttpClient, {
      siteUrl: this.context.pageContext.web.absoluteUrl,
      sourceLibraryTitle: libraryTitle,
      archiveFolderServerRelativeUrl: archiveFolder
    });

    // OUTBOUND: Moves selected files to archive folder; INBOUND: per-item move results.
    const results = await lifecycleService.moveItemsToArchive(itemIds);
    const failed = results.filter((r) => !r.moved);

    if (failed.length > 0) {
      const details = failed.map((f) => `ID ${f.itemId}: ${f.error}`).join('; ');
      throw new Error(`Archive completed with ${failed.length} failure(s): ${details}`);
    }

    await Dialog.alert(`Archived ${results.length} document(s).`);
  }

  private _parseNumericIds(items: readonly ISelectedItemPayload[]): number[] {
    return items
      .map((item) => Number(item.id))
      .filter((id) => Number.isFinite(id) && id > 0);
  }

  private _updateCommandVisibility(selectedCount: number): void {
    // INBOUND: Extension properties controlling visible actions and selection requirement.
    const mode = this.properties.actionMode || 'flow';
    const requiresSelection = this.properties.requireSelection !== false;
    const hasValidSelection = !requiresSelection || selectedCount > 0;
    const showSelectionActions = mode === 'flow' || mode === 'both' || mode === 'all';

    // OUTBOUND: Sets `.visible` on SharePoint command bar button definitions.
    if (this._flowCommand) {
      this._flowCommand.visible = showSelectionActions && hasValidSelection;
    }

    if (this._batchUpdateCommand) {
      this._batchUpdateCommand.visible = showSelectionActions && hasValidSelection;
    }

    if (this._teamsMessageCommand) {
      this._teamsMessageCommand.visible = showSelectionActions && hasValidSelection;
    }

    if (this._archiveDocumentsCommand) {
      this._archiveDocumentsCommand.visible = showSelectionActions && hasValidSelection;
    }

    if (this._pbiCommand) {
      this._pbiCommand.visible = mode === 'powerbi' || mode === 'both' || mode === 'all';
    }
  }

  private async _triggerPowerAutomateFlow(
    webhookUrl: string,
    selectedItems: ISelectedItemPayload[]
  ): Promise<void> {
    // INBOUND: Current user, site, and list metadata from SharePoint pageContext + mapped items.
    const payload = buildFlowTriggerPayload({
      triggeredAtUtc: new Date().toISOString(),
      user: {
        displayName: this.context.pageContext.user.displayName,
        email: this.context.pageContext.user.email,
        loginName: this.context.pageContext.user.loginName
      },
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      listId: this.context.pageContext.list?.id?.toString(),
      listTitle: this.context.pageContext.list?.title,
      selectedItems
    });

    // OUTBOUND: Serialized JSON request body and headers sent to Power Automate.
    const requestOptions: IHttpClientOptions = {
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    };

    // OUTBOUND: HTTP POST to webhookUrl; INBOUND: HttpClientResponse from Power Automate.
    const response: HttpClientResponse = await this.context.httpClient.post(
      webhookUrl,
      HttpClient.configurations.v1,
      requestOptions
    );

    if (!response.ok) {
      // INBOUND: Error response body/status from Power Automate endpoint.
      const errorText = await response.text();
      // OUTBOUND: Throws formatted error back to onExecute error handler.
      throw new Error(`Flow HTTP ${response.status}: ${errorText || response.statusText}`);
    }
  }
}