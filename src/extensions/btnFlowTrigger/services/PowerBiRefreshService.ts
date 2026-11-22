import { AadHttpClientFactory } from '@microsoft/sp-http';
import {
  AadHttpClient,
  HttpClientResponse,
  IHttpClientOptions
} from '@microsoft/sp-http';

export type RefreshNotifyOption = 'MailOnFailure' | 'NoNotification';

export interface IPowerBiRefreshConfig {
  workspaceId?: string;
  datasetId?: string;
  apiScope?: string;
  notifyOption?: RefreshNotifyOption;
}

const DEFAULT_PBI_SCOPE: string = 'https://analysis.windows.net/powerbi/api/.default';

export class PowerBiRefreshService {
  private readonly _aadHttpClientFactory: AadHttpClientFactory;
  private readonly _workspaceId?: string;
  private readonly _datasetId?: string;
  private readonly _resourceScope: string;
  private readonly _notifyOption: RefreshNotifyOption;
  private _aadClient?: AadHttpClient;

  // INBOUND: SPFx AadHttpClientFactory and caller configuration (workspaceId, datasetId, apiScope, notifyOption).
  public constructor(
    aadHttpClientFactory: AadHttpClientFactory,
    config: IPowerBiRefreshConfig
  ) {
    this._aadHttpClientFactory = aadHttpClientFactory;
    this._workspaceId = config.workspaceId;
    this._datasetId = config.datasetId;
    // Strip trailing "/.default" because AadHttpClientFactory expects the base resource URI.
    this._resourceScope = (config.apiScope || DEFAULT_PBI_SCOPE).replace(/\/\.default$/, '');
    this._notifyOption = config.notifyOption || 'MailOnFailure';
  }

  // Returns true if both workspaceId and datasetId are set; used as a config guard before refreshing.
  public isConfigured(): boolean {
    return Boolean(this._workspaceId && this._datasetId);
  }

  public async initialize(): Promise<void> {
    if (!this._aadClient) {
      // OUTBOUND: Requests Entra ID-authenticated client for _resourceScope.
      // INBOUND: Resolves to an initialized AadHttpClient instance.
      this._aadClient = await this._aadHttpClientFactory.getClient(this._resourceScope);
    }
  }

  /**
 * Triggers an asynchronous Power BI dataset refresh via the Power BI REST API and returns the HTTP status code.
 * Uses AadHttpClient so SPFx handles Entra ID token acquisition for the configured Power BI scope, and explicitly
 * accepts 202 Accepted because Power BI queues dataset refreshes asynchronously rather than completing synchronously.
 */
  public async triggerRefresh(): Promise<number> {
    if (!this._workspaceId || !this._datasetId) {
      throw new Error('Power BI workspaceId and datasetId must be configured.');
    }

    if (!this._aadClient) {
      await this.initialize();
    }

    if (!this._aadClient) {
      throw new Error('AadHttpClient is not initialized. Verify Entra ID API permission grants.');
    }

    const endpoint = `https://api.powerbi.com/v1.0/myorg/groups/${encodeURIComponent(
      this._workspaceId
    )}/datasets/${encodeURIComponent(this._datasetId)}/refreshes`;

    const requestOptions: IHttpClientOptions = {
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ notifyOption: this._notifyOption })
    };

    // OUTBOUND: POST request to Power BI refresh endpoint.
    // INBOUND: HttpClientResponse from Power BI.
    const response: HttpClientResponse = await this._aadClient.post(
      endpoint,
      AadHttpClient.configurations.v1,
      requestOptions
    );

    // Power BI queues refreshes asynchronously and commonly returns 202 Accepted.
    if (!response.ok && response.status !== 202) {
      // INBOUND: Error body text from Power BI API.
      const errorText = await response.text();
      throw new Error(`Power BI API HTTP ${response.status}: ${errorText || response.statusText}`);
    }

    return response.status;
  } 
}