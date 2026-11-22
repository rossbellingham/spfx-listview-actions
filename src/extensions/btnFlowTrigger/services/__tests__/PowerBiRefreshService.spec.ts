jest.mock(
  '@microsoft/sp-http',
  () => ({
    AadHttpClient: {
      configurations: { v1: 'v1' }
    }
  }),
  { virtual: true }
);

import { AadHttpClient } from '@microsoft/sp-http';
import { PowerBiRefreshService } from '../PowerBiRefreshService';

describe('PowerBiRefreshService', () => {
  it('initializes client and returns 202 Accepted on refresh', async () => {
    const postMock = jest.fn().mockResolvedValue({
      ok: false,
      status: 202,
      statusText: 'Accepted'
    });
    const mockClient = { post: postMock } as unknown as AadHttpClient;
    const provider = { getClient: jest.fn().mockResolvedValue(mockClient) };

    const service = new PowerBiRefreshService(
      { workspaceId: 'ws-123', datasetId: 'ds-456' },
      provider
    );

    const status = await service.triggerRefresh();

    expect(provider.getClient).toHaveBeenCalledWith('https://analysis.windows.net/powerbi/api');
    expect(status).toBe(202);
  });

  it('throws when workspaceId or datasetId is missing', async () => {
    const service = new PowerBiRefreshService({ workspaceId: '', datasetId: 'ds-456' });

    await expect(service.triggerRefresh()).rejects.toThrow(
      /workspaceId and datasetId must be configured/i
    );
  });
});