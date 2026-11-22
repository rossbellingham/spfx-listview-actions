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
import { GraphActionService } from '../GraphActionService';

describe('GraphActionService', () => {
  const createMockAadClient = (postMock: jest.Mock): AadHttpClient =>
    ({ post: postMock } as unknown as AadHttpClient);

  it('posts a channel message and returns the Graph message ID', async () => {
    const postMock = jest.fn().mockResolvedValue({
      ok: true,
      status: 201,
      json: jest.fn().mockResolvedValue({ id: 'msg-12345' })
    });

    const service = new GraphActionService(createMockAadClient(postMock));
    const messageId = await service.postChannelMessage({
      teamId: 'team-1',
      channelId: '19:channel@thread.tacv2',
      bodyHtml: '<p>Item flagged for review</p>'
    });

    expect(messageId).toBe('msg-12345');
    expect(postMock).toHaveBeenCalledWith(
      'https://graph.microsoft.com/v1.0/teams/team-1/channels/19%3Achannel%40thread.tacv2/messages',
      'v1',
      expect.objectContaining({
        body: JSON.stringify({
          body: {
            contentType: 'html',
            content: '<p>Item flagged for review</p>'
          }
        })
      })
    );
  });

  it('throws descriptive error when Graph returns non-OK status', async () => {
    const postMock = jest.fn().mockResolvedValue({
      ok: false,
      status: 403,
      statusText: 'Forbidden',
      text: jest.fn().mockResolvedValue('Missing ChannelMessage.Send permission')
    });

    const service = new GraphActionService(createMockAadClient(postMock));

    await expect(
      service.postChannelMessage({
        teamId: 'team-1',
        channelId: 'chan-1',
        bodyHtml: '<p>Test</p>'
      })
    ).rejects.toThrow(/Microsoft Graph HTTP 403: Missing ChannelMessage\.Send permission/);
  });

  it('throws when required message inputs are missing', async () => {
    const postMock = jest.fn();
    const service = new GraphActionService(createMockAadClient(postMock));

    await expect(
      service.postChannelMessage({ teamId: '', channelId: 'chan-1', bodyHtml: '<p>Hi</p>' })
    ).rejects.toThrow(/teamId, channelId, and bodyHtml are required/i);
    expect(postMock).not.toHaveBeenCalled();
  });
});