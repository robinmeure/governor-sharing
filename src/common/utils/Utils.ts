

import { IFacepilePersona } from '@fluentui/react';
import { SharePointIdentitySet } from '@microsoft/microsoft-graph-types';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPaginationFilterState } from '../../webparts/sharing/components/SharingList/SharingDetailedList';


// need to rework this sorting method to be a) working with dates and b) be case insensitive
export function genericSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}
// thanks to Michael Norward for this function, https://stackoverflow.com/questions/8900732/sort-objects-in-an-array-alphabetically-on-one-property-of-the-array
// export function textSort(objectsArr: any[], prop, isSortedDescending = true): any[] {
//   // eslint-disable-next-line no-prototype-builtins
//   const objectsHaveProp = objectsArr.every((object: any) => object.hasOwnProperty(prop));
//   if (objectsHaveProp) {
//     const newObjectsArr = objectsArr.slice();
//     newObjectsArr.sort((a, b) => {
//       if (isNaN(Number(a[prop]))) {
//         const textA = a[prop].toUpperCase(),
//           textB = b[prop].toUpperCase();
//         if (isSortedDescending) {
//           return textA < textB ? -1 : textA > textB ? 1 : 0;
//         } else {
//           return textB < textA ? -1 : textB > textA ? 1 : 0;
//         }
//       } else {
//         return isSortedDescending ? a[prop] - b[prop] : b[prop] - a[prop];
//       }
//     });
//     return newObjectsArr;
//   }
//   return objectsArr;
// }


// export function convertUserToFacePilePersona(identity: SharePointIdentitySet): IFacepilePersona {
//   if (identity.siteUser) {
//     const siteUser = identity.siteUser;
//     const _user: IFacepilePersona =
//     {
//       data: (siteUser.loginName && siteUser.loginName.indexOf('#ext') !== -1) ? "Guest" : "Member",
//       personaName: siteUser.displayName ?? undefined,
//       name: siteUser.loginName ? siteUser.loginName.replace("i:0#.f|membership|", "") : undefined
//     };
//     return _user;
//   }
//   else if (identity.siteGroup) {
//     const siteGroup = identity.siteGroup;
//     const _user: IFacepilePersona =
//     {
//       data: "Group",
//       personaName: siteGroup.displayName ?? undefined,
//       name: siteGroup.loginName ? siteGroup.loginName.replace("c:0t.c|tenant|", "") : undefined
//     };
//     return _user;
//   }
//   else {
//     const _user: IFacepilePersona =
//     {
//       name: identity.user?.id ?? undefined,
//       data: (identity.user?.id === null) ? "Guest" : "Member",
//       personaName: identity.user?.displayName ?? undefined
//     };
//     return _user;
//   }
// }

// export function convertToFacePilePersona(identities: SharePointIdentitySet[]): IFacepilePersona[] {
//   const _users: IFacepilePersona[] = [];
//   if (identities.length > 1) {
//     identities.forEach((user) => {
//       if (user.siteUser) {
//         const siteUser = user.siteUser;
//         const _user: IFacepilePersona =
//         {
//           data: (siteUser.loginName && siteUser.loginName.indexOf('#ext') !== -1) ? "Guest" : "Member",
//           personaName: siteUser.displayName ?? undefined,
//           name: siteUser.loginName ? siteUser.loginName.replace("i:0#.f|membership|", "") : undefined
//         };
//         _users.push(_user);
//       }
//       else {
//         const _user: IFacepilePersona =
//         {
//           name: user.user?.id ?? undefined,
//           data: (user.user?.id === null) ? "Guest" : "Member",
//           personaName: user.user?.displayName ?? undefined
//         };
//         _users.push(_user);
//       }
//     });
//   }
//   else if (identities.length === 1) {
//     _users.push(convertUserToFacePilePersona(identities[0]));
//   }

//   return _users;
// }

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


export const searchQueryGeneratorForDocs = (context: WebPartContext, queryFilter: IPaginationFilterState): string => {
  const tenantId = context.pageContext.aadInfo.tenantId;
  const everyoneExceptExternalsUserName = `spo-grid-all-users/${tenantId}`;

  const filterVal = queryFilter.filterVal;
  const searchQuery = queryFilter.searchQuery ? queryFilter.searchQuery + " " : "";
  const siteFilter = filterVal.siteUrl ? `(SPSiteUrl:${filterVal.siteUrl}) ` : "";
  const testFilter = DEBUG ? "" : "";
  let fileFolderFilter = "(IsDocument:TRUE OR IsContainer:TRUE) ";
  if (filterVal.fileFolder === "OnlyFiles") {
    fileFolderFilter = "(IsDocument:TRUE OR IsContainer:FALSE) ";
  } else if (filterVal.fileFolder === "OnlyFolders") {
    fileFolderFilter = "(IsDocument:FALSE OR IsContainer:TRUE) ";
  }

  let query = `${testFilter}${searchQuery}${siteFilter}${fileFolderFilter} AND (NOT FileExtension:aspx) AND ((SharedWithUsersOWSUSER:*) OR (SharedWithUsersOWSUSER:${everyoneExceptExternalsUserName} OR SharedWithUsersOWSUser:Everyone))`;

  // let query = `${testFilter}${searchQuery}${siteFilter}${fileFolderFilter} AND (NOT FileExtension:aspx)`;

  let isTeams = false, isPrivateChannel = false;
  let groupId = "";
  if (context.sdks.microsoftTeams) {
    isTeams = true;
  }
  if (isTeams) {
    isPrivateChannel = context.sdks.microsoftTeams && (context.sdks.microsoftTeams.context.channelType === "Private") || false;
    groupId = context.sdks.microsoftTeams?.context?.groupId ?? "";
    // const siteUrl = context.sdks.microsoftTeams ? context.sdks.microsoftTeams.context.teamSiteUrl : '';
    if (!isPrivateChannel)
      query = `(IsDocument:TRUE OR IsContainer:TRUE) AND (NOT FileExtension:aspx) AND ((SharedWithUsersOWSUSER:*) OR (SharedWithUsersOWSUSER:${everyoneExceptExternalsUserName} OR SharedWithUsersOWSUser:Everyone)) AND (GroupId:${groupId} OR RelatedGroupId:${groupId})`;
  }

  return query;
}





