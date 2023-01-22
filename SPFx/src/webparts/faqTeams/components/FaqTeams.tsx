/* eslint-disable @typescript-eslint/typedef */
/* eslint-disable @typescript-eslint/no-unused-vars */
import * as React from "react";
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import styles from "./FaqTeams.module.scss";
import { IFaqTeamsProps } from "./IFaqTeamsProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { App } from "./App";
import "../../../ExternalRef/CSS/Style.css";
import { Web } from "@pnp/sp/webs";

export default class FaqTeams extends React.Component<IFaqTeamsProps, {}> {
  // eslint-disable-next-line @typescript-eslint/explicit-member-accessibility

  public render(): React.ReactElement<IFaqTeamsProps> {
    const TeamName = this.props.context.sdks.microsoftTeams.context.teamName;
    const ChannelName =
      this.props.context.sdks.microsoftTeams.context.channelName;
    // const TeamName = "";
    // const ChannelName = "";
    return (
      <App
        tenantURL={this.props.tenantURL}
        siteName={this.props.siteName}
        teamName={TeamName}
        channelName={ChannelName}
        // teamsContext={TeamsContext}
        // userContext={UserContext}
      />
    );
  }
}
