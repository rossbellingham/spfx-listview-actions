import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import {
  HttpClient,
  HttpClientResponse,
  IHttpClientOptions
} from '@microsoft/sp-http';
import * as strings from 'BtnFlowTriggerCommandSetStrings';
import { PowerBiRefreshService } from './services/PowerBiRefreshService';

export interface IBtnFlowTriggerCommandSetProperties {
  actionMode?: 'flow' | 'powerbi' | 'both';
  flowWebhookUrl?: string;
  powerBiWorkspaceId?: string;
  powerBiDatasetId?: string;
  powerBiApiScope?: string;
  requireSelection?: boolean;
}

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

const LOG_SOURCE: string = 'BtnFlowTriggerCommandSet';
const CMD_TRIGGER_FLOW: string = 'CMD_TRIGGER_FLOW';
const CMD_REFRESH_PBI: string = 'CMD_REFRESH_PBI';

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
  private _powerBiService?: PowerBiRefreshService;

  // Runs once when the ListView Command Set is initialized ("onload" setup).
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initializing BtnFlowTriggerCommandSet');

    this._flowCommand = this.tryGetCommand(CMD_TRIGGER_FLOW);
    this._pbiCommand = this.tryGetCommand(CMD_REFRESH_PBI);

    // INBOUND: Extension properties from manifest, serve.json, or SharePoint deployment config.
    // OUTBOUND: Passing Power BI target config + AadHttpClientFactory into PowerBiRefreshService.
    this._powerBiService = new PowerBiRefreshService(this.context.aadHttpClientFactory, {
      workspaceId: this.properties.powerBiWorkspaceId,
      datasetId: this.properties.powerBiDatasetId,
      apiScope: this.properties.powerBiApiScope
    });

    try {
      // OUTBOUND: Requests AAD client initialization for Power BI API scope.
      await this._powerBiService.initialize();
    } catch (err) {
      // INBOUND: Initialization error caught so command set still loads; retry happens on execute.
      Log.warn(LOG_SOURCE, `PowerBiRefreshService init deferred: ${String(err)}`);
    }

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);
    // INBOUND: Initial selected row count from SharePoint list view state.
    this._updateCommandVisibility(this.context.listView.selectedRows?.length ?? 0);
    return Promise.resolve();
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
    const selectedItems = mapSelectedRows(event.selectedRows || []);

    try {
      switch (event.itemId) {
        case CMD_TRIGGER_FLOW: {
          if (!this.properties.flowWebhookUrl) {
            await Dialog.alert(strings.MissingConfigMessage);
            return;
          }
          // OUTBOUND: Sends selectedItems + SharePoint context to Power Automate HTTP webhook.
          await this._triggerPowerAutomateFlow(this.properties.flowWebhookUrl, selectedItems);
          await Dialog.alert(strings.FlowSuccessMessage.replace('{0}', String(selectedItems.length)));
          break;
        }

        case CMD_REFRESH_PBI: {
          if (!this._powerBiService || !this._powerBiService.isConfigured()) {
            await Dialog.alert(strings.MissingConfigMessage);
            return;
          }
          // OUTBOUND: Invokes Power BI dataset refresh; INBOUND: receives HTTP status code (e.g. 202).
          const statusCode = await this._powerBiService.triggerRefresh();
          await Dialog.alert(strings.PbiRefreshAcceptedMessage.replace('{0}', String(statusCode)));
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

  private _updateCommandVisibility(selectedCount: number): void {
    // INBOUND: Extension properties controlling visible actions and selection requirement.
    const mode = this.properties.actionMode || 'flow';
    const requiresSelection = this.properties.requireSelection !== false;
    const hasValidSelection = !requiresSelection || selectedCount > 0;

    // OUTBOUND: Sets `.visible` on SharePoint command bar button definitions.
    if (this._flowCommand) {
      this._flowCommand.visible = (mode === 'flow' || mode === 'both') && hasValidSelection;
    }

    if (this._pbiCommand) {
      this._pbiCommand.visible = mode === 'powerbi' || mode === 'both';
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