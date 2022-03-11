import * as React from 'react';
import { IPnPMgtSpFxDemoProps } from './IPnPMgtSpFxDemoProps';
import { Icon, Pivot, PivotItem } from 'office-ui-fabric-react';
import { Agenda, Person, FileList, Tasks, Todo, Get, MgtTemplateProps } from '@microsoft/mgt-react/dist/es6/spfx';
import { avatarType, PersonCardInteraction, ViewType } from '@microsoft/mgt-spfx';
import styles from './PnPMgtSpFxDemo.module.scss';

const getMailScopes = ["Mail.Read"];
const getTeamScopes = ["Team.ReadBasic.All"];

const MailTemplate = (props: MgtTemplateProps) => {
  const mail = props.dataContext;
  return (
    <div>
      <h3>{mail.subject}</h3>
      <h4>
        <Person
          personQuery={mail.sender.emailAddress.address}
          view={ViewType.oneline}
          personCardInteraction={PersonCardInteraction.hover}>
        </Person>
      </h4>
      <p>{mail.bodyPreview}</p>
    </div>
  );
};

const TeamTemplate = (props: MgtTemplateProps) => {
  const team = props.dataContext;
  return (
    <div>
      <h3>
        <Icon iconName="TeamsLogo" className={styles.teamsIcon}></Icon>
        {team.displayName}
      </h3>
      <p>{team.description}</p>
    </div>
  );
};

export default class PnPMgtSpFxDemo extends React.Component<IPnPMgtSpFxDemoProps, {}> {
  public render(): React.ReactElement<IPnPMgtSpFxDemoProps> {
    const {
      description
    } = this.props;

    return (
      <div>
        <h1>{description}</h1>
        <Person
          personQuery="me"
          view={ViewType.threelines}
          fetchImage={true}
          avatarType={avatarType.photo}
          personCardInteraction={PersonCardInteraction.hover}>
        </Person>
        <Pivot>
          <PivotItem headerText="Agenda">
            <Agenda groupByDay={true} showMax={7}></Agenda>
          </PivotItem>
          <PivotItem headerText="Files">
            <FileList enableFileUpload={true}></FileList>
          </PivotItem>
          <PivotItem headerText="Tasks">
            <Tasks></Tasks>
          </PivotItem>
          <PivotItem headerText="To Do">
            <Todo></Todo>
          </PivotItem>
          <PivotItem headerText="Mails">
            <Get resource="/me/messages" version="v1.0" scopes={getMailScopes} maxPages={2}>
              <MailTemplate template="value" />
            </Get>
          </PivotItem>
          <PivotItem headerText="Teams">
            <Get resource="/me/joinedTeams" version="v1.0" scopes={getTeamScopes} maxPages={2}>
              <TeamTemplate template="value" />
            </Get>
          </PivotItem>
        </Pivot>
      </div>
    );
  }
}
