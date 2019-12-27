import * as React from "react";
import { SPUser } from "../sharepoint/spt.sharepoint.entities";
import { Constants } from "../spt.constants";

export interface ISitePermGroupsProps {
    user: SPUser;
}

export class SitePermissionsGroups extends React.Component<ISitePermGroupsProps, {}> {
    constructor(props: ISitePermGroupsProps) {
        super(props);
    }

    render() {
        return <div id="userPermissionGroups">
            <br />
            {
                !this.props.user.Groups &&
                <div>{Constants.getLiteral("directoryNoGroups")}</div>
            }
            {
                this.props.user.Groups && this.props.user.Groups.map((group) =>
                    <div>{group.Name}</div>
                )
            }
            <br />
        </div>;
    }
}
