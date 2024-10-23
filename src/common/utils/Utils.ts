/* eslint-disable */

import { IColumn, IContextualMenuItem, IFacepilePersona } from '@fluentui/react';
import { IdentitySet, Permission, SearchRequest, SearchResponse } from '@microsoft/microsoft-graph-types';
import { isEqual } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISearchResultExtended } from '../../webparts/sharing/components/SharingView/ISearchResultExtended';
import { _CONST } from './Const';
import { ISharingResult } from '../../webparts/sharing/model';

// need to rework this sorting method to be a) working with dates and b) be case insensitive
export function genericSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  const key = columnKey as keyof T;
  return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}
// thanks to Michael Norward for this function, https://stackoverflow.com/questions/8900732/sort-objects-in-an-array-alphabetically-on-one-property-of-the-array
export function textSort(objectsArr: any[], prop, isSortedDescending = true): any[] {
  // eslint-disable-next-line no-prototype-builtins
  const objectsHaveProp = objectsArr.every((object: any) => object.hasOwnProperty(prop));
  if (objectsHaveProp) {
    const newObjectsArr = objectsArr.slice();
    newObjectsArr.sort((a, b) => {
      if (isNaN(Number(a[prop]))) {
        const textA = a[prop].toUpperCase(),
          textB = b[prop].toUpperCase();
        if (isSortedDescending) {
          return textA < textB ? -1 : textA > textB ? 1 : 0;
        } else {
          return textB < textA ? -1 : textB > textA ? 1 : 0;
        }
      } else {
        return isSortedDescending ? a[prop] - b[prop] : b[prop] - a[prop];
      }
    });
    return newObjectsArr;
  }
  return objectsArr;
}

export function uniqForObject<T>(array: T[]): T[] {
  const result: T[] = [];
  for (const item of array) {
    const found = result.some((value) => isEqual(value, item));
    if (!found) {
      result.push(item);
    }
  }
  return result;
}

export function rightTrim(sourceString: string, searchString: string): string {
  for (; ;) {
    const pos = sourceString.lastIndexOf(searchString);
    if (pos === sourceString.length - 1) {
      const result = sourceString.slice(0, pos);
      sourceString = result;
    }
    else {
      break;
    }
  }
  return sourceString;
}

export function convertToGraphUserFromLinkKind(linkKind: number): microsoftgraph.User {
  const _user: microsoftgraph.User = {};
  switch (linkKind) {
    case 2: _user.displayName = "Organization View"; break;
    case 3: _user.displayName = "Organization Edit"; break;
    case 4: _user.displayName = "Anonymous View"; break;
    case 5: _user.displayName = "Anonymous Edit"; break;
    default: break;
  }
  _user.userType = "Link";
  return _user;
}

export function convertUserToFacePilePersona(user: IdentitySet): IFacepilePersona {
  if (user["siteUser"]) {
    const siteUser = user["siteUser"];
    const _user: IFacepilePersona =
    {
      data: (siteUser.loginName.indexOf('#ext') !== -1) ? "Guest" : "Member",
      personaName: siteUser.displayName,
      name: siteUser.loginName.replace("i:0#.f|membership|", "")
    };
    return _user;
  }
  else if (user["siteGroup"]) {
    const siteGroup = user["siteGroup"];
    const _user: IFacepilePersona =
    {
      data: "Group",
      personaName: siteGroup.displayName,
      name: siteGroup.loginName.replace("c:0t.c|tenant|", "")
    };
    return _user;
  }
  else {
    const _user: IFacepilePersona =
    {
      name: user.user.id,
      data: (user.user.id === null) ? "Guest" : "Member",
      personaName: user.user.displayName
    };
    return _user;
  }

}

export function convertToFacePilePersona(users: IdentitySet[]): IFacepilePersona[] {
  const _users: IFacepilePersona[] = [];
  if (users.length > 1) {
    users.forEach((user) => {
      if (user["siteUser"] !== null) {
        const siteUser = user["siteUser"];
        const _user: IFacepilePersona =
        {
          data: (siteUser.loginName.indexOf('#ext') !== -1) ? "Guest" : "Member",
          personaName: siteUser.displayName,
          name: siteUser.loginName.replace("i:0#.f|membership|", "")
        };
        _users.push(_user);
      }
      else {
        const _user: IFacepilePersona =
        {
          name: user.user.id,
          data: (user.user.id === null) ? "Guest" : "Member",
          personaName: user.user.displayName
        };
        _users.push(_user);
      }
    });
  }
  else {
    _users.push(convertUserToFacePilePersona(users[0]));
  }

  return _users;
}

export function getSortingMenuItems(column: IColumn, onSortColumn: (column: IColumn, isSortedDescending: boolean) => void): IContextualMenuItem[] {
  const menuItems = [];
  if (column.data === Number) {
    menuItems.push(
      {
        key: 'smallToLarger',
        name: 'Smaller to larger',
        canCheck: true,
        checked: column.isSorted && !column.isSortedDescending,
        onClick: () => onSortColumn(column, false)
      },
      {
        key: 'largerToSmall',
        name: 'Larger to smaller',
        canCheck: true,
        checked: column.isSorted && column.isSortedDescending,
        onClick: () => onSortColumn(column, true)
      }
    );
  }
  else if (column.data === Date) {
    menuItems.push(
      {
        key: 'oldToNew',
        name: 'Older to newer',
        canCheck: true,
        checked: column.isSorted && !column.isSortedDescending,
        onClick: () => onSortColumn(column, false)
      },
      {
        key: 'newToOld',
        name: 'Newer to Older',
        canCheck: true,
        checked: column.isSorted && column.isSortedDescending,
        onClick: () => onSortColumn(column, true)
      }
    );
  }
  else
  //(column.data == String) 
  // NOTE: in case of 'complex columns like Taxonomy, you need to add more logic'
  {
    menuItems.push(
      {
        key: 'aToZ',
        name: 'A to Z',
        canCheck: true,
        checked: column.isSorted && !column.isSortedDescending,
        onClick: () => onSortColumn(column, false)
      },
      {
        key: 'zToA',
        name: 'Z to A',
        canCheck: true,
        checked: column.isSorted && column.isSortedDescending,
        onClick: () => onSortColumn(column, true)
      }
    );
  }
  return menuItems;
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


export const SearchQueryGeneratorForDocs = (context: WebPartContext): string => {
  try {
    const tenantId = context.pageContext.aadInfo.tenantId;
    const everyoneExceptExternalsUserName = `spo-grid-all-users/${tenantId}`;
    // let query = `(IsDocument:TRUE OR IsContainer:TRUE) AND (NOT FileExtension:aspx) AND ((SharedWithUsersOWSUSER:*) OR (SharedWithUsersOWSUSER:${everyoneExceptExternalsUserName} OR SharedWithUsersOWSUser:Everyone))`;
    let query = `(IsDocument:TRUE OR IsContainer:TRUE) AND (NOT FileExtension:aspx)`;

    let siteUrl = context.pageContext.web.absoluteUrl;
    let isTeams: boolean, isPrivateChannel = false;
    let groupId = "";
    if (context.sdks.microsoftTeams) {
      isTeams = true;
    }
    if (isTeams) {
      isPrivateChannel = (context.sdks.microsoftTeams.context.channelType === "Private");
      groupId = context.sdks.microsoftTeams.context.groupId;
      siteUrl = context.sdks.microsoftTeams.context.teamSiteUrl;
      if (!isPrivateChannel)
        query = `(IsDocument:TRUE OR IsContainer:TRUE) AND (NOT FileExtension:aspx) AND ((SharedWithUsersOWSUSER:*) OR (SharedWithUsersOWSUSER:${everyoneExceptExternalsUserName} OR SharedWithUsersOWSUser:Everyone)) AND (GroupId:${groupId} OR RelatedGroupId:${groupId})`;
    }

    return query;
  } catch (error) {
    console.log("FazLog ~ SearchQueryGeneratorForDocs ~ error:", error);
    throw error;
  }
}


export const GraphResponseToSearchResultMapper = (searchResponse: SearchResponse[]): ISearchResultExtended[] => {

  try {
    const locSearchResultExtended: ISearchResultExtended[] = [];
    searchResponse.forEach(results => {
      results.hitsContainers.forEach(hits => {
        hits?.hits?.forEach((hit: any) => {
          const SharedWithUsersOWSUser = (hit.resource.listItem.fields.sharedWithUsersOWSUSER !== undefined) ? hit.resource.listItem.fields.sharedWithUsersOWSUSER : null;

          // if we don't get a driveId back (e.g. documentlibrary), then skip the returned item
          if (hit.resource.listItem.fields.driveId === undefined)
            return;

          const result: ISearchResultExtended = {
            DriveItemId: hit.resource.id,
            FileName: hit.resource.listItem.fields.fileName,
            FileExtension: hit.resource.listItem.fields.fileExtension ? hit.resource.listItem.fields.fileExtension : "folder",
            ListId: hit.resource.listItem.fields.listId,
            FileId: hit.resource.listItem.id,
            DriveId: hit.resource.listItem.fields.driveId,
            ListItemId: hit.resource.listItem.fields.listItemId,
            Path: hit.resource.webUrl,
            LastModifiedTime: hit.resource.lastModifiedDateTime,
            SharedWithUsersOWSUSER: SharedWithUsersOWSUser,
            SiteUrl: hit.resource.listItem.fields.spSiteUrl
          }
          locSearchResultExtended.push(result);
        });
      });
    });
    return locSearchResultExtended;
  } catch (error) {
    console.log("FazLog ~ SearchResultMapper ~ error:", error);
    throw error;
  }
}

export const SearchResultAndDriveItemToSharingMapper = (file: ISearchResultExtended, driveItem: Permission, standardGroups: string[]): ISharingResult => {
  console.log("FazLog ~ SearchResultAndDriveItemToSharingMapper ~ file:", file);

  try {
    let sharedWithUser: IFacepilePersona[] = [];
    let sharingUserType = "Member";

    // Getting all the details of the file and in which folder is lives
    let folderUrl = file.Path.replace(`/${file.FileName}`, '');
    let folderName = folderUrl.lastIndexOf("/") > 0 ? folderUrl.substring(folderUrl.lastIndexOf("/") + 1) : folderUrl;

    // for certain filetypes we get the dispform.aspx link back instead of the full path, so we need to fix that
    if (folderName.indexOf("DispForm.aspx") > -1) {
      folderUrl = folderUrl.substring(0, folderUrl.lastIndexOf("/Forms/DispForm.aspx"));
      folderName = folderUrl.lastIndexOf("/") > 0 ? folderUrl.substring(folderUrl.lastIndexOf("/") + 1) : folderUrl;
      file.FileExtension = file.FileName.substring(file.FileName.lastIndexOf(".") + 1);
    }


    if (driveItem.link) {
      switch (driveItem.link.scope) {
        case "anonymous":
          break;
        case "organization": {
          const _user: IFacepilePersona = {};
          _user.personaName = driveItem.link.scope + " " + driveItem.link.type;
          _user.data = "Organization";
          if (sharedWithUser.indexOf(_user) === -1) {
            sharedWithUser.push(_user);
          }
          break;
        }
        case "users": {
          const _users = convertToFacePilePersona(driveItem.grantedToIdentitiesV2);
          sharedWithUser.push(..._users);
          break;
        }
        default:
          break;
      }
    }
    else // checking the normal permissions as well, other than the sharing links
    {
      // if the permission is not the same as the default associated spo groups, we need to add it to the sharedWithUser array
      if (standardGroups.indexOf(driveItem.grantedTo.user.displayName) === -1) {
        const _users = convertUserToFacePilePersona(driveItem.grantedToV2);
        sharedWithUser.push(_users);
      }
      else // otherwise, we're gonna add these groups and mark it as inherited permissions
      {
        const _user: IFacepilePersona = {};
        _user.personaName = driveItem.grantedTo.user.displayName;
        _user.data = "Inherited";
        if (sharedWithUser.indexOf(_user) === -1) {
          sharedWithUser.push(_user);
        }
      }
    }

    if (file.SharedWithUsersOWSUSER !== null) {
      const _users = processUsers(file.SharedWithUsersOWSUSER);
      sharedWithUser.push(..._users);
    }

    // if there are any duplicates, this will remove them (e.g. multiple organization links)
    sharedWithUser = uniqForObject(sharedWithUser);
    //TODO check
    // if (sharedWithUser.length === 0)
    //     continue;


    let isGuest = false;
    let isLink = false;
    let isInherited = false;

    for (const user of sharedWithUser) {
      switch (user.data) {
        case "Guest": isGuest = true; break;
        case "Organization": isLink = true; break;
        case "Inherited": isInherited = true; break;
      }
    }

    // if we found a guest user, we need to set the sharingUserType to Guest
    if (isGuest) {
      sharingUserType = "Guest";
    }
    else if (isLink) {
      sharingUserType = "Link";
    }
    else if (isInherited) {
      sharingUserType = "Inherited";
    }

    // building up the result to be returned
    const sharedResult: ISharingResult =
    {
      FileExtension: file?.FileExtension ? file.FileExtension : "folder",
      FileName: file.FileName,
      Channel: folderName,
      LastModified: file.LastModifiedTime,
      SharedWith: sharedWithUser,
      ListId: file.ListId,
      ListItemId: file.ListItemId,
      Url: file.Path,
      FolderUrl: folderUrl,
      SharingUserType: sharingUserType,
      FileId: file.FileId,
      SiteUrl: file.SiteUrl
    };
    return sharedResult;
  } catch (error) {
    console.log("FazLog ~ SearchResultToSharingMapper ~ error:", error);
    throw error;
  }
}