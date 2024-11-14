

import { IFacepilePersona } from '@fluentui/react';
import { SharePointIdentitySet } from '@microsoft/microsoft-graph-types';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPaginationFilterState } from '../../webparts/sharing/components/SharingList/SharingDetailedList';


export function convertUserToFacePilePersona(identity: SharePointIdentitySet): IFacepilePersona {
  if (identity.siteUser) {
    const siteUser = identity.siteUser;
    const _user: IFacepilePersona =
    {
      data: (siteUser.loginName && siteUser.loginName.indexOf('#ext') !== -1) ? "Guest" : "Member",
      personaName: siteUser.displayName ?? undefined,
      name: siteUser.loginName ? siteUser.loginName.replace("i:0#.f|membership|", "") : undefined
    };
    return _user;
  }
  else if (identity.siteGroup) {
    const siteGroup = identity.siteGroup;
    const _user: IFacepilePersona =
    {
      data: "Group",
      personaName: siteGroup.displayName ?? undefined,
      name: siteGroup.loginName ? siteGroup.loginName.replace("c:0t.c|tenant|", "") : undefined
    };
    return _user;
  }
  else {
    const _user: IFacepilePersona =
    {
      name: identity.user?.id ?? undefined,
      data: (identity.user?.id === null) ? "Guest" : "Member",
      personaName: identity.user?.displayName ?? undefined
    };
    return _user;
  }
}

export function convertToFacePilePersona(identities: SharePointIdentitySet[]): IFacepilePersona[] {
  const _users: IFacepilePersona[] = [];
  if (identities.length > 1) {
    identities.forEach((user) => {
      if (user.siteUser) {
        const siteUser = user.siteUser;
        const _user: IFacepilePersona =
        {
          data: (siteUser.loginName && siteUser.loginName.indexOf('#ext') !== -1) ? "Guest" : "Member",
          personaName: siteUser.displayName ?? undefined,
          name: siteUser.loginName ? siteUser.loginName.replace("i:0#.f|membership|", "") : undefined
        };
        _users.push(_user);
      }
      else {
        const _user: IFacepilePersona =
        {
          name: user.user?.id ?? undefined,
          data: (user.user?.id === null) ? "Guest" : "Member",
          personaName: user.user?.displayName ?? undefined
        };
        _users.push(_user);
      }
    });
  }
  else if (identities.length === 1) {
    _users.push(convertUserToFacePilePersona(identities[0]));
  }

  return _users;
}

/// this is used to process the SharedWithUsersOWSUSER output to get the userPrincipalName and userType
export function processUsers(users: string): IFacepilePersona[] {
  const _users: microsoftgraph.User[] = [];

  if (users === null || users === undefined)
    return _users;

  if (users.match("\n\n")) {
    const allUsers = users.split("\n\n");
    allUsers.forEach(element => {
      const user: IFacepilePersona = {
        personaName: element.split("|")[1].trim(),
        data: (element.indexOf("#ext#") > -1) ? "Guest" : "Member",
        id: element.split("|")[0].trim()
      };
      _users.push(user)
    });
  }
  else {
    const user: IFacepilePersona = {
      personaName: users.split("|")[1].trim(),
      data: (users.indexOf("#ext#") > -1) ? "Guest" : "Member",
      id: users.split("|")[0].trim()
    };
    _users.push(user)
  }
  return _users;
}


export const searchQueryGeneratorForFiles = (context: WebPartContext, queryFilter: IPaginationFilterState): string => {
  const filterVal = queryFilter.filterVal;
  const searchQuery = queryFilter.searchKeyword ? queryFilter.searchKeyword + " " : "";
  const siteFilter = filterVal.siteUrl ? `(SPSiteUrl:${filterVal.siteUrl}) ` : "";
  const testFilter = DEBUG ? "" : "";
  const preQuery = filterVal.preQuery ? `${filterVal.preQuery} ` : "";

  let fileFolderFilter = "(IsDocument:TRUE OR IsContainer:TRUE) ";
  if (filterVal.fileFolder === "OnlyFiles") {
    fileFolderFilter = "(IsDocument:TRUE OR IsContainer:FALSE) ";
  } else if (filterVal.fileFolder === "OnlyFolders") {
    fileFolderFilter = "(IsDocument:FALSE OR IsContainer:TRUE) ";
  }

  let query = `${preQuery}${testFilter}${searchQuery}${siteFilter}${fileFolderFilter} AND (NOT FileExtension:aspx)`;
  //SharedWithUsersOWSUSER not working with all the items
  // AND ((SharedWithUsersOWSUSER:*) OR SharedWithUsersOWSUser:Everyone))`;

  let isTeams = false, isPrivateChannel = false;
  let groupId = "";
  if (context.sdks.microsoftTeams) {
    isTeams = true;
  }
  if (isTeams) {
    isPrivateChannel = context.sdks.microsoftTeams && (context.sdks.microsoftTeams.context.channelType === "Private") || false;
    groupId = context.sdks.microsoftTeams?.context?.groupId ?? "";
    if (!isPrivateChannel)
      // query = `(IsDocument:TRUE OR IsContainer:TRUE) AND (NOT FileExtension:aspx) AND ((SharedWithUsersOWSUSER:*) OR (SharedWithUsersOWSUSER:${everyoneExceptExternalsUserName} OR SharedWithUsersOWSUser:Everyone)) AND (GroupId:${groupId} OR RelatedGroupId:${groupId})`;
      query = `(IsDocument:TRUE OR IsContainer:TRUE) AND (NOT FileExtension:aspx) AND (GroupId:${groupId} OR RelatedGroupId:${groupId})`;
  }

  return query;
}





