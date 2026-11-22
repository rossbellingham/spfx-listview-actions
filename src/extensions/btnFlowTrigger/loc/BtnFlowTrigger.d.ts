declare interface IBtnFlowTriggerCommandSetStrings {
  MissingConfigMessage: string;
  FlowSuccessMessage: string;
  PbiRefreshAcceptedMessage: string;
  ActionFailedMessage: string;
  CMD_TRIGGER_FLOW_TITLE: string;
  CMD_REFRESH_POWERBI_TITLE: string;
  CMD_BATCH_UPDATE_TITLE: string;
  CMD_POST_TEAMS_MESSAGE_TITLE: string;
  CMD_ARCHIVE_DOCUMENTS_TITLE: string;
}

declare module 'BtnFlowTriggerCommandSetStrings' {
  const strings: IBtnFlowTriggerCommandSetStrings;
  export = strings;
}