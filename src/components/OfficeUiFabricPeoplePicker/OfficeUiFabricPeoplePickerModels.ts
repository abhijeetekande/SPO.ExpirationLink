import { IPersonaProps, IPersona } from "office-ui-fabric-react";

export const enum TypePicker {
    Normal = "Normal",
    Compact = "Compact"
}

export const enum PrincipalType {
    None                = 0,
    User                = 1 << 0,
    DistributionList    = 1 << 1,
    SecurityGroup       = 1 << 2,
    SharePointGroup     = 1 << 3,
}      

export interface IOfficeUiFabricPeoplePickerState {
    selectedItems: IPersonaProps[];
}

export interface IClientPeoplePickerSearchUser {
    Description: string;
    DisplayText: string;
    EntityData: {
        IsAltSecIdPresent: string;
        ObjectId: string;
        Title: string;
        Email: string;
        MobilePhone: string;
        OtherMails: string;
        Department: string;
    };
    EntityType: string;
    IsResolved: boolean;
    Key: string;
    MultipleMatches: any[];
    ProviderDisplayName: string;
    ProviderName: string;
}

export interface IUserListItem {
    Id: number;
    Title: string;
}

export interface IEnsureUser extends IUserListItem {
    Email: string;
    IsEmailAuthenticationGuestUser: boolean;
    IsHiddenInUI: boolean;
    IsShareByEmailGuestUser: boolean;
    IsSiteAdmin: boolean;
    LoginName: string;
    PrincipalType: number;
    UserId: {
        NameId: string;
        NameIdIssuer: string;
    };
}

export interface ISPDataUserInfoItem extends IUserListItem {
    IsSiteAdmin: boolean;
    Name: string;
    UserName: string;
    Department?: string;
    FirstName?: string;
    LastName?: string;
    JobTitle?: string;
}

export interface IEnsurableSharePointUser 
    extends IClientPeoplePickerSearchUser, IEnsureUser {}

export interface IUserPersona extends IPersonaProps {
    user: IUserListItem;
}

export const SharePointUserInfoPersona = (user: ISPDataUserInfoItem) => ({
    user,
    primaryText: user.Title,
    secondaryText: user.JobTitle,
    tertiaryText: user.Department,
    imageShouldFadeIn: true,
    imageUrl: `/_layouts/15/userphoto.aspx?size=S&accountname=${user.UserName}`
} as IUserPersona);

export const SharePointSearchUserPersona = (user: IEnsurableSharePointUser) => ({
    user,
    primaryText: user.DisplayText,
    secondaryText: user.EntityData.Title,
    tertiaryText: user.EntityData.Department,
    imageShouldFadeIn: true,
    imageUrl: `/_layouts/15/userphoto.aspx?size=S&accountname=${user.Key.substr(user.Key.lastIndexOf('|') + 1)}`
} as IUserPersona);
