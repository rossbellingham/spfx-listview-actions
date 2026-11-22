import { AadHttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

export type PowerBiNotifyOption = 'NoNotification' | 'MailOnFailure' | 'MailOnCompletion';

export interface IPowerBiRefreshConfig {
  workspaceId: string;
  datasetId: string;
  notifyOption?: PowerBiNotifyOption;
  resourceUri?: string;
}

export interface IAadClientProvider {
  getClient: (resourceUri: string) => Promise<AadHttpClient>;
}

export class PowerBiRefreshService {
  private readonly _aadClientProvider?: IAadClientProvider;
  private readonly _resourceUri: string;
  private readonly _workspaceId: string;
  private readonly _datasetId: string;
  private readonly _notifyOption: PowerBiNotifyOption;
  private _aadClient?: AadHttpClient;

  constructor(config: IPowerBiRefreshConfig, aadClientProvider?: IAadClientProvider, aadClient?: AadHttpClient) {
    this._workspaceId = config.workspaceId;
    this._datasetId = config.datasetId;
    this._notifyOption = config.notifyOption ?? 'NoNotification';
    this._resourceUri = config.resourceUri ?? 'https://analysis.windows.net/powerbi/api';
    this._aadClientProvider = aadClientProvider;
    this._aadClient = aadClient;
  }

  public isConfigured(): boolean {
    return Boolean(this._workspaceId && this._datasetId);
  }

  public async initialize(): Promise<void> {
    if (!this._aadClientProvider) {
      return;
    }

    // OUTBOUND: Requests Entra ID-scoped AadHttpClient for Power BI API.
    // INBOUND: Initialized AadHttpClient instance.
    this._aadClient = await this._aadClientProvider.getClient(this._resourceUri);
  }

  /**
   * Triggers an asynchronous Power BI dataset refresh via the Power BI REST API and returns the HTTP status code.
   * Uses AadHttpClient so SPFx handles Entra ID token acquisition for the Power BI scope, and explicitly
   * accepts 202 Accepted because Power BI queues dataset refreshes asynchronously.
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

    if (!response.ok && response.status !== 202) {
      // INBOUND: Error response body text from Power BI REST API.
      const errorText = await response.text();
      throw new Error(`Power BI API HTTP ${response.status}: ${errorText || response.statusText}`);
    }

    return response.status;
  }
}