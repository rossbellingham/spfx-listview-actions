import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

export interface IBatchUpdateConfig {
  siteUrl: string;
  listTitle: string;
  batchSize?: number;
}

export interface IBatchItemResult {
  itemId: number;
  success: boolean;
  error?: string;
}

export interface IBatchUpdateSummary {
  total: number;
  succeeded: number;
  failed: number;
  results: IBatchItemResult[];
}

export class SharePointBatchUpdateService {
  private readonly _spHttpClient: SPHttpClient;
  private readonly _siteUrl: string;
  private readonly _listTitle: string;
  private readonly _batchSize: number;

  constructor(spHttpClient: SPHttpClient, config: IBatchUpdateConfig) {
    this._spHttpClient = spHttpClient;
    this._siteUrl = config.siteUrl.replace(/\/+$/, '');
    this._listTitle = config.listTitle;
    this._batchSize = config.batchSize && config.batchSize > 0 ? config.batchSize : 10;
  }

  public isConfigured(): boolean {
    return Boolean(this._spHttpClient && this._siteUrl && this._listTitle);
  }

  /**
   * Updates field values across multiple list or library items in capped batches and returns per-item results.
   * Splitting large selections into smaller batches helps avoid overwhelming SharePoint at once, while capturing
   * errors per item stops a single bad item from cancelling the rest of the update.
   */
  public async updateItems(
    itemIds: number[],
    fields: Record<string, unknown>
  ): Promise<IBatchUpdateSummary> {
    if (!this.isConfigured()) {
      throw new Error('SharePointBatchUpdateService requires spHttpClient, siteUrl, and listTitle.');
    }

    if (!itemIds || itemIds.length === 0) {
      return { total: 0, succeeded: 0, failed: 0, results: [] };
    }

    const results: IBatchItemResult[] = [];
    const encodedTitle = encodeURIComponent(this._listTitle);
    const requestOptions: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify(fields)
    };

    for (let i = 0; i < itemIds.length; i += this._batchSize) {
      const chunk = itemIds.slice(i, i + this._batchSize);

      const chunkResults = await Promise.all(
        chunk.map(async (itemId): Promise<IBatchItemResult> => {
          const endpoint = `${this._siteUrl}/_api/web/lists/getbytitle('${encodedTitle}')/items(${itemId})`;

          try {
            // OUTBOUND: PATCH/MERGE update request for a list item via SPHttpClient.
            // INBOUND: SPHttpClientResponse from SharePoint REST API.
            const response: SPHttpClientResponse = await this._spHttpClient.post(
              endpoint,
              SPHttpClient.configurations.v1,
              requestOptions
            );

            if (!response.ok) {
              // INBOUND: Error response body text from SharePoint REST API.
              const errorText = await response.text();
              return {
                itemId,
                success: false,
                error: `HTTP ${response.status}: ${errorText || response.statusText}`
              };
            }

            return { itemId, success: true };
          } catch (err) {
            return {
              itemId,
              success: false,
              error: err instanceof Error ? err.message : String(err)
            };
          }
        })
      );

      results.push(...chunkResults);
    }

    const succeeded = results.filter((r) => r.success).length;
    return {
      total: results.length,
      succeeded,
      failed: results.length - succeeded,
      results
    };
  }
}