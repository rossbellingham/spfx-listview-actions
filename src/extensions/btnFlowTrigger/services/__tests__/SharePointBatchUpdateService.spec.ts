jest.mock(
  '@microsoft/sp-http',
  () => ({
    SPHttpClient: {
      configurations: { v1: 'v1' }
    }
  }),
  { virtual: true }
);

import { SPHttpClient } from '@microsoft/sp-http';
import { SharePointBatchUpdateService } from '../SharePointBatchUpdateService';

describe('SharePointBatchUpdateService', () => {
  const createMockClient = (handler: jest.Mock): SPHttpClient =>
    ({ post: handler } as unknown as SPHttpClient);

  it('returns empty summary when no item IDs are provided', async () => {
    const postMock = jest.fn();
    const service = new SharePointBatchUpdateService(createMockClient(postMock), {
      siteUrl: 'https://contoso.sharepoint.com/sites/ops',
      listTitle: 'Tasks'
    });

    const summary = await service.updateItems([], { Status: 'Approved' });

    expect(summary).toEqual({ total: 0, succeeded: 0, failed: 0, results: [] });
    expect(postMock).not.toHaveBeenCalled();
  });

  it('throws when service configuration is incomplete', async () => {
    const postMock = jest.fn();
    const service = new SharePointBatchUpdateService(createMockClient(postMock), {
      siteUrl: '',
      listTitle: 'Tasks'
    });

    await expect(service.updateItems([1], { Status: 'Approved' })).rejects.toThrow(
      /requires spHttpClient, siteUrl, and listTitle/i
    );
  });

  it('updates items in chunks and records partial failures without aborting', async () => {
    const postMock = jest
      .fn()
      .mockResolvedValueOnce({ ok: true, status: 204 })
      .mockResolvedValueOnce({
        ok: false,
        status: 400,
        statusText: 'Bad Request',
        text: jest.fn().mockResolvedValue('Invalid field value')
      })
      .mockResolvedValueOnce({ ok: true, status: 204 });

    const service = new SharePointBatchUpdateService(createMockClient(postMock), {
      siteUrl: 'https://contoso.sharepoint.com/sites/ops/',
      listTitle: 'Tasks',
      batchSize: 2
    });

    const summary = await service.updateItems([10, 20, 30], { Status: 'Done' });

    expect(postMock).toHaveBeenCalledTimes(3);
    expect(postMock).toHaveBeenNthCalledWith(
      1,
      "https://contoso.sharepoint.com/sites/ops/_api/web/lists/getbytitle('Tasks')/items(10)",
      'v1',
      expect.objectContaining({
        body: JSON.stringify({ Status: 'Done' })
      })
    );
    expect(summary).toEqual({
      total: 3,
      succeeded: 2,
      failed: 1,
      results: [
        { itemId: 10, success: true },
        { itemId: 20, success: false, error: 'HTTP 400: Invalid field value' },
        { itemId: 30, success: true }
      ]
    });
  });
});