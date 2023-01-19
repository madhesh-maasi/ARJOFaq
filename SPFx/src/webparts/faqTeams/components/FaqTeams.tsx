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
    const list = "FAQ";
    const TeamName = this.props.context.sdks.microsoftTeams.context.teamName;
    const ChannelName =
      this.props.context.sdks.microsoftTeams.context.channelName;
    // const TeamName = "";
    // const ChannelName = "";
    return (
      <App
        siteUrl={this.props.siteUrl}
        list={list}
        teamName={TeamName}
        channelName={ChannelName}
        // teamsContext={TeamsContext}
        // userContext={UserContext}
      />
    );
  }
}
