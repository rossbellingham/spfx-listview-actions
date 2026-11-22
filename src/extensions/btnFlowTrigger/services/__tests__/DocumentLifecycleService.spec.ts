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
import { DocumentLifecycleService } from '../DocumentLifecycleService';

describe('DocumentLifecycleService', () => {
  const createMockClient = (getMock: jest.Mock, postMock: jest.Mock): SPHttpClient =>
    ({ get: getMock, post: postMock } as unknown as SPHttpClient);

  it('fetches file metadata and moves file to archive folder', async () => {
    const getMock = jest.fn().mockResolvedValue({
      ok: true,
      json: jest.fn().mockResolvedValue({
        FileLeafRef: 'Q4 Report.docx',
        FileRef: '/sites/ops/Shared Documents/Q4 Report.docx'
      })
    });
    const postMock = jest.fn().mockResolvedValue({ ok: true, status: 200 });

    const service = new DocumentLifecycleService(createMockClient(getMock, postMock), {
      siteUrl: 'https://contoso.sharepoint.com/sites/ops',
      sourceLibraryTitle: 'Documents',
      archiveFolderServerRelativeUrl: '/sites/ops/Shared Documents/Archive/'
    });

    const results = await service.moveItemsToArchive([42]);

    expect(results).toEqual([{ itemId: 42, moved: true }]);
    expect(postMock).toHaveBeenCalledWith(
      "https://contoso.sharepoint.com/sites/ops/_api/web/getfilebyserverrelativeurl('%2Fsites%2Fops%2FShared%20Documents%2FQ4%20Report.docx')/MoveTo(newUrl='%2Fsites%2Fops%2FShared%20Documents%2FArchive%2FQ4%20Report.docx',flags=1)",
      'v1',
      expect.any(Object)
    );
  });

  it('reports non-file items without calling MoveTo', async () => {
    const getMock = jest.fn().mockResolvedValue({
      ok: true,
      json: jest.fn().mockResolvedValue({})
    });
    const postMock = jest.fn();

    const service = new DocumentLifecycleService(createMockClient(getMock, postMock), {
      siteUrl: 'https://contoso.sharepoint.com/sites/ops',
      sourceLibraryTitle: 'Documents',
      archiveFolderServerRelativeUrl: '/sites/ops/Shared Documents/Archive'
    });

    const results = await service.moveItemsToArchive([99]);

    expect(results).toEqual([
      {
        itemId: 99,
        moved: false,
        error: 'Item does not expose FileLeafRef and FileRef (not a file item).'
      }
    ]);
    expect(postMock).not.toHaveBeenCalled();
  });
});