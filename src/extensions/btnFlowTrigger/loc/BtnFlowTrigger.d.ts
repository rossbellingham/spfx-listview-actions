declare interface IBtnFlowTriggerCommandSetStrings {
  FlowCommandTitle: string;
  RefreshPbiCommandTitle: string;
  MissingConfigMessage: string;
  FlowSuccessMessage: string;
  PbiRefreshAcceptedMessage: string;
  ActionFailedMessage: string;
}

declare module 'BtnFlowTriggerCommandSetStrings' {
  const strings: IBtnFlowTriggerCommandSetStrings;
  export = strings;
}
