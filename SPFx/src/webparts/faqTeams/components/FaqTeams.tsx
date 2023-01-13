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
    return <App siteUrl={this.props.siteUrl} />;
  }
}
