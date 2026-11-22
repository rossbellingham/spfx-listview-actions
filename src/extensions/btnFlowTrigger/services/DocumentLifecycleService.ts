import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

export interface IDocumentLifecycleConfig {
  siteUrl: string;
  sourceLibraryTitle: string;
  archiveFolderServerRelativeUrl: string;
}

export interface IFileMoveResult {
  itemId: number;
  moved: boolean;
  error?: string;
}

export class DocumentLifecycleService {
  private readonly _spHttpClient: SPHttpClient;
  private readonly _siteUrl: string;
  private readonly _sourceLibraryTitle: string;
  private readonly _archiveFolderServerRelativeUrl: string;

  constructor(spHttpClient: SPHttpClient, config: IDocumentLifecycleConfig) {
    this._spHttpClient = spHttpClient;
    this._siteUrl = config.siteUrl.replace(/\/+$/, '');
    this._sourceLibraryTitle = config.sourceLibraryTitle;
    this._archiveFolderServerRelativeUrl = config.archiveFolderServerRelativeUrl;
  }

  public isConfigured(): boolean {
    return Boolean(
      this._spHttpClient &&
        this._siteUrl &&
        this._sourceLibraryTitle &&
        this._archiveFolderServerRelativeUrl
    );
  }

  /**
   * Moves selected library items by ID to a configured archive folder via SharePoint REST "MoveTo".
   * Resolves each item's "FileLeafRef" and "FileRef" first so destination paths stay accurate even if a user
   * selects nested items or files with spaces in their names.
   */
  public async moveItemsToArchive(itemIds: number[]): Promise<IFileMoveResult[]> {
    if (!this.isConfigured()) {
      throw new Error(
        'DocumentLifecycleService requires spHttpClient, siteUrl, sourceLibraryTitle, and archiveFolderServerRelativeUrl.'
      );
    }

    if (!itemIds || itemIds.length === 0) {
      return [];
    }

    const encodedLibrary = encodeURIComponent(this._sourceLibraryTitle);
    const normalizedArchiveFolder = this._archiveFolderServerRelativeUrl.replace(/\/+$/, '');
    const results: IFileMoveResult[] = [];

    for (const itemId of itemIds) {
      const itemEndpoint =
        `${this._siteUrl}/_api/web/lists/getbytitle('${encodedLibrary}')/items(${itemId})` +
        `?$select=FileLeafRef,FileRef`;

      try {
        // OUTBOUND: GET file metadata (FileLeafRef, FileRef) for item before moving.
        // INBOUND: SPHttpClientResponse containing file path metadata.
        const metaResponse: SPHttpClientResponse = await this._spHttpClient.get(
          itemEndpoint,
          SPHttpClient.configurations.v1
        );

        if (!metaResponse.ok) {
          // INBOUND: Error text from SharePoint metadata lookup.
          const errText = await metaResponse.text();
          results.push({
            itemId,
            moved: false,
            error: `Metadata lookup HTTP ${metaResponse.status}: ${errText || metaResponse.statusText}`
          });
          continue;
        }

        // INBOUND: Parsed file metadata fields.
        const meta = (await metaResponse.json()) as {
          FileLeafRef?: string;
          FileRef?: string;
        };

        if (!meta.FileLeafRef || !meta.FileRef) {
          results.push({
            itemId,
            moved: false,
            error: 'Item does not expose FileLeafRef and FileRef (not a file item).'
          });
          continue;
        }

        const destinationServerRelativeUrl = `${normalizedArchiveFolder}/${meta.FileLeafRef}`;
        const moveEndpoint =
          `${this._siteUrl}/_api/web/getfilebyserverrelativeurl('${encodeURIComponent(meta.FileRef)}')` +
          `/MoveTo(newUrl='${encodeURIComponent(destinationServerRelativeUrl)}',flags=1)`;

        const moveOptions: ISPHttpClientOptions = {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        };

        // OUTBOUND: POST MoveTo request to relocate file to archive folder.
        // INBOUND: SPHttpClientResponse from SharePoint MoveTo operation.
        const moveResponse: SPHttpClientResponse = await this._spHttpClient.post(
          moveEndpoint,
          SPHttpClient.configurations.v1,
          moveOptions
        );

        if (!moveResponse.ok) {
          // INBOUND: Error text from SharePoint MoveTo endpoint.
          const moveErr = await moveResponse.text();
          results.push({
            itemId,
            moved: false,
            error: `MoveTo HTTP ${moveResponse.status}: ${moveErr || moveResponse.statusText}`
          });
          continue;
        }

        results.push({ itemId, moved: true });
      } catch (err) {
        results.push({
          itemId,
          moved: false,
          error: err instanceof Error ? err.message : String(err)
        });
      }
    }

    return results;
  }
}