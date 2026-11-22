import { AadHttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';

export interface IGraphActionConfig {
  graphBaseUrl?: string;
}

export interface ITeamsChannelMessageInput {
  teamId: string;
  channelId: string;
  bodyHtml: string;
}

export class GraphActionService {
  private readonly _aadClient: AadHttpClient;
  private readonly _graphBaseUrl: string;

  constructor(aadClient: AadHttpClient, config: IGraphActionConfig = {}) {
    this._aadClient = aadClient;
    this._graphBaseUrl = (config.graphBaseUrl || 'https://graph.microsoft.com/v1.0').replace(
      /\/+$/,
      ''
    );
  }

  public isConfigured(): boolean {
    return Boolean(this._aadClient && this._graphBaseUrl);
  }

  /**
   * Posts an HTML message to a Microsoft Teams channel via Microsoft Graph and returns the created message ID.
   * Uses AadHttpClient so SPFx handles Entra ID token acquisition against Microsoft Graph rather than
   * embedding tokens or routing simple notifications through an extra Power Automate flow.
   */
  public async postChannelMessage(input: ITeamsChannelMessageInput): Promise<string> {
    if (!this.isConfigured()) {
      throw new Error('GraphActionService is not initialized with an AadHttpClient.');
    }

    if (!input.teamId || !input.channelId || !input.bodyHtml) {
      throw new Error('teamId, channelId, and bodyHtml are required to post a Teams message.');
    }

    const endpoint = `${this._graphBaseUrl}/teams/${encodeURIComponent(
      input.teamId
    )}/channels/${encodeURIComponent(input.channelId)}/messages`;

    const requestOptions: IHttpClientOptions = {
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        body: {
          contentType: 'html',
          content: input.bodyHtml
        }
      })
    };

    // OUTBOUND: POST message payload to Microsoft Graph Teams channel endpoint.
    // INBOUND: HttpClientResponse from Microsoft Graph.
    const response: HttpClientResponse = await this._aadClient.post(
      endpoint,
      AadHttpClient.configurations.v1,
      requestOptions
    );

    if (!response.ok) {
      // INBOUND: Error response text from Microsoft Graph API.
      const errorText = await response.text();
      throw new Error(`Microsoft Graph HTTP ${response.status}: ${errorText || response.statusText}`);
    }

    // INBOUND: Created message JSON payload from Microsoft Graph.
    const data = (await response.json()) as { id?: string };
    return data.id ?? '';
  }
}