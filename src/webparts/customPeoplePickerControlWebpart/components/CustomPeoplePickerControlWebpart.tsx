import * as React from 'react';
import styles from './CustomPeoplePickerControlWebpart.module.scss';
import { ICustomPeoplePickerControlWebpartProps } from './ICustomPeoplePickerControlWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { PeoplePicker, PrincipalType, PeopleFilterRuleMode } from '../../../controls/PeoplePicker';

export default class CustomPeoplePickerControlWebpart extends React.Component<ICustomPeoplePickerControlWebpartProps, {}> {

  private _filterFunction = (res) => {
    return !(res && res.EntityData && res.EntityData.PrincipalType && res.EntityData.PrincipalType === "GUEST_USER");
  }

  public render(): React.ReactElement<ICustomPeoplePickerControlWebpartProps> {

    return (
      <div className={ styles.customPeoplePickerControlWebpart }>

        {/* OPTION 1: HIDE EXTERNAL USERS BOOLEAN */}
        <PeoplePicker context={this.props.context}
                      titleText="Option 1: People Picker using Hide External users Flag"
                      personSelectionLimit={5}
                      principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup, PrincipalType.SecurityGroup, PrincipalType.DistributionList]}
                      suggestionsLimit={5}
                      resolveDelay={200}

                      hideExternalUsers={true}
          />

        {/* OPTION 2: FILTER CONFIGURATION OBJECT */}
        <PeoplePicker context={this.props.context}
                      titleText="Option 2: People Picker using Filter Configuration Object"
                      personSelectionLimit={5}
                      principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup, PrincipalType.SecurityGroup, PrincipalType.DistributionList]}
                      suggestionsLimit={5}
                      resolveDelay={200}

                      filterConfig={{
                        hideExternalUsers: true
                      }}
          />

        {/* OPTION 3: FILTER RULES COLLECTION */}
        <PeoplePicker context={this.props.context}
                      titleText="Option 3: People Picker using Filter Rules Collection"
                      personSelectionLimit={5}
                      principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup, PrincipalType.SecurityGroup, PrincipalType.DistributionList]}
                      suggestionsLimit={5}
                      resolveDelay={200}

                      filterRules={[
                        { property: "PrincipalType", value: "GUEST_USER" },
                        // { property: "Department", value: "Information Technology" }
                        // { property: "Title", value: "Wizard" }
                      ]}
                      filterRulesMode={PeopleFilterRuleMode.Exclude}
          />

        {/* OPTION 4: FILTER FUNCTION */}
        <PeoplePicker context={this.props.context}
                      titleText="Option 4: People Picker using Filter Function"
                      personSelectionLimit={5}
                      principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup, PrincipalType.SecurityGroup, PrincipalType.DistributionList]}
                      suggestionsLimit={5}
                      resolveDelay={200}

                      filterFunction={this._filterFunction}
          />

      </div>
    );
  }
}
