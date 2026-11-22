jest.mock('@microsoft/sp-http', () => ({
  AadHttpClient: {
    configurations: {
      v1: {}
    }
  }
}));

jest.mock('@microsoft/sp-core-library', () => ({
  Log: {
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn()
  }
}));

jest.mock('@microsoft/sp-listview-extensibility', () => ({
  BaseListViewCommandSet: class {}
}));

jest.mock('@microsoft/sp-dialog', () => ({
  Dialog: {
    alert: jest.fn().mockResolvedValue(undefined)
  }
}));

jest.mock(
  'BtnFlowTriggerCommandSetStrings',
  () => ({
    MissingConfigMessage: 'Missing configuration',
    FlowSuccessMessage: 'Triggered {0} items',
    PbiRefreshAcceptedMessage: 'Refresh status: {0}',
    ActionFailedMessage: 'Failed: {0}'
  }),
  { virtual: true }
);

import { AadHttpClient, AadHttpClientFactory, HttpClientResponse } from '@microsoft/sp-http';
import { PowerBiRefreshService } from './PowerBiRefreshService';
import { buildFlowTriggerPayload, IListRowReader, mapSelectedRows } from '../BtnFlowTriggerCommandSet';

describe('PowerBiRefreshService', () => {
  const mockPost = jest.fn();
  const mockGetClient = jest.fn();

  const mockAadClient = {
    post: mockPost
  } as unknown as AadHttpClient;

  const mockFactory = {
    getClient: mockGetClient
  } as unknown as AadHttpClientFactory;

  const workspaceId = '11111111-1111-1111-1111-111111111111';
  const datasetId = '22222222-2222-2222-2222-222222222222';

  beforeEach(() => {
    jest.clearAllMocks();
    mockGetClient.mockResolvedValue(mockAadClient);
  });

  describe('isConfigured()', () => {
    it('returns false when workspaceId is missing', () => {
      const service = new PowerBiRefreshService(mockFactory, { datasetId });
      expect(service.isConfigured()).toBe(false);
    });

    it('returns false when datasetId is missing', () => {
      const service = new PowerBiRefreshService(mockFactory, { workspaceId });
      expect(service.isConfigured()).toBe(false);
    });

    it('returns true when both workspaceId and datasetId are provided', () => {
      const service = new PowerBiRefreshService(mockFactory, { workspaceId, datasetId });
      expect(service.isConfigured()).toBe(true);
    });
  });

  describe('initialize() and triggerRefresh()', () => {
    it('normalizes /.default scope and returns 202 when refresh is accepted', async () => {
      mockPost.mockResolvedValue({
        ok: false,
        status: 202,
        statusText: 'Accepted',
        text: jest.fn().mockResolvedValue('')
      } as unknown as HttpClientResponse);

      const service = new PowerBiRefreshService(mockFactory, {
        workspaceId,
        datasetId,
        apiScope: 'https://analysis.windows.net/powerbi/api/.default'
      });

      await service.initialize();
      const status = await service.triggerRefresh();

      expect(mockGetClient).toHaveBeenCalledTimes(1);
      expect(mockGetClient).toHaveBeenCalledWith('https://analysis.windows.net/powerbi/api');
      expect(mockPost).toHaveBeenCalledWith(
        `https://api.powerbi.com/v1.0/myorg/groups/${workspaceId}/datasets/${datasetId}/refreshes`,
        AadHttpClient.configurations.v1,
        expect.objectContaining({
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ notifyOption: 'MailOnFailure' })
        })
      );
      expect(status).toBe(202);
    });

    it('lazily initializes AAD client if triggerRefresh() is called before initialize()', async () => {
      mockPost.mockResolvedValue({
        ok: true,
        status: 200,
        statusText: 'OK',
        text: jest.fn().mockResolvedValue('')
      } as unknown as HttpClientResponse);

      const service = new PowerBiRefreshService(mockFactory, { workspaceId, datasetId });
      const status = await service.triggerRefresh();

      expect(mockGetClient).toHaveBeenCalledTimes(1);
      expect(status).toBe(200);
    });

    it('throws when triggerRefresh() is called without workspaceId or datasetId', async () => {
      const service = new PowerBiRefreshService(mockFactory, { workspaceId });

      await expect(service.triggerRefresh()).rejects.toThrow(
        'Power BI workspaceId and datasetId must be configured.'
      );
      expect(mockPost).not.toHaveBeenCalled();
    });

    it('throws with HTTP status and response body when Power BI returns an error', async () => {
      mockPost.mockResolvedValue({
        ok: false,
        status: 403,
        statusText: 'Forbidden',
        text: jest.fn().mockResolvedValue('Insufficient privileges')
      } as unknown as HttpClientResponse);

      const service = new PowerBiRefreshService(mockFactory, { workspaceId, datasetId });

      await expect(service.triggerRefresh()).rejects.toThrow(
        'Power BI API HTTP 403: Insufficient privileges'
      );
    });

    it('falls back to statusText when error response body is empty', async () => {
      mockPost.mockResolvedValue({
        ok: false,
        status: 503,
        statusText: 'Service Unavailable',
        text: jest.fn().mockResolvedValue('')
      } as unknown as HttpClientResponse);

      const service = new PowerBiRefreshService(mockFactory, { workspaceId, datasetId });

      await expect(service.triggerRefresh()).rejects.toThrow(
        'Power BI API HTTP 503: Service Unavailable'
      );
    });
  });
});

describe('BtnFlowTriggerCommandSet payload helpers', () => {
  it('maps list rows to normalized items and falls back to FileLeafRef when Title is missing', () => {
    const rows: IListRowReader[] = [
      {
        getValueByName: (name: string) => {
          const values: Record<string, unknown> = {
            ID: 42,
            Title: 'Q4 Budget',
            FileLeafRef: 'q4-budget.xlsx',
            ContentType: 'Document'
          };
          return values[name];
        }
      },
      {
        getValueByName: (name: string) => {
          const values: Record<string, unknown> = {
            ID: 7,
            Title: null,
            FileLeafRef: 'notes.docx'
          };
          return values[name];
        }
      }
    ];

    expect(mapSelectedRows(rows)).toEqual([
      {
        id: '42',
        title: 'Q4 Budget',
        fileLeafRef: 'q4-budget.xlsx',
        contentType: 'Document'
      },
      {
        id: '7',
        title: 'notes.docx',
        fileLeafRef: 'notes.docx',
        contentType: undefined
      }
    ]);
  });

  it('builds the Power Automate payload with user, SharePoint context, and selected items', () => {
    const payload = buildFlowTriggerPayload({
      triggeredAtUtc: '2026-11-22T16:00:00.000Z',
      user: {
        displayName: 'Ross Bellingham',
        email: 'ross@example.com',
        loginName: 'i:0#.f|membership|ross@example.com'
      },
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/ops',
      listId: 'list-guid-123',
      listTitle: 'Requests',
      selectedItems: [{ id: '1', title: 'Item One' }]
    });

    expect(payload).toEqual({
      triggeredAtUtc: '2026-11-22T16:00:00.000Z',
      triggeredBy: {
        displayName: 'Ross Bellingham',
        email: 'ross@example.com',
        loginName: 'i:0#.f|membership|ross@example.com'
      },
      sharePointContext: {
        webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/ops',
        listId: 'list-guid-123',
        listTitle: 'Requests'
      },
      selectedItems: [{ id: '1', title: 'Item One' }]
    });
  });
});