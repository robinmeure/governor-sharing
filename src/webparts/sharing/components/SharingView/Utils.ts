
import { IContextualMenuItem, IColumn,IFacepilePersona } from '@fluentui/react';
import ISharingResult from './ISharingResult';
import { isEqual } from '@microsoft/sp-lodash-subset';
import { IdentitySet } from '@microsoft/microsoft-graph-types';

export class Utils 
{
    // need to rework this sorting method to be a) working with dates and b) be case insensitive
    public GenericSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] 
    {
        const key = columnKey as keyof T;
        return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
    }
    // thanks to Michael Norward for this function, https://stackoverflow.com/questions/8900732/sort-objects-in-an-array-alphabetically-on-one-property-of-the-array
    public TextSort(objectsArr, prop, isSortedDescending = true) 
    {
        let objectsHaveProp = objectsArr.every(object => object.hasOwnProperty(prop));
        if(objectsHaveProp)    {
            let newObjectsArr = objectsArr.slice();
            newObjectsArr.sort((a, b) => {
                if(isNaN(Number(a[prop])))  {
                    let textA = a[prop].toUpperCase(),
                        textB = b[prop].toUpperCase();
                    if(isSortedDescending)   {
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

    public uniqForObject<T>(array: T[]): T[] {
        const result: T[] = [];
        for (const item of array) {
            const found = result.some((value) => isEqual(value, item));
            if (!found) {
                result.push(item);
            }
        }
        return result;
    }
    // this method is used to group the results by the sharedWith property
    // creating a single row for each document and having a facepile object of sharedWith users
    // public processGroupedResults(results:any) :  ISharingResult[]
    // {
    //     const _sharedWithSpecified:ISharingResult[] = [];
    //     Object.keys(results).forEach((key) => 
    //     {
    //         let files = results[key];
    //         let sharedWithUsers:IFacepilePersona[] = [];
    //         let isGuestOrEveryone:boolean = false;
    //         let sharingType:string;    

    //         if (files.length == 1)
    //         {
    //             _sharedWithSpecified.push(files[0]);
    //         }
    //         else
    //         {
    //             files.forEach((sharedResult) =>
    //             {
    //                 const facepilePersona:IFacepilePersona = { personaName : sharedResult.SharedWith.displayName}
    //                 sharedWithUsers.push(facepilePersona)
    //                 // if the user is either guest or everyone, then we need to show the sharing type as guest or everyone. This for showing the proper icon in front of the document
    //                 if (sharedResult.SharingUserType == "Guest" || sharedResult.SharingUserType == "Everyone" 
    //                     || (sharedResult.SharingUserType == "Link" && sharedResult.SharedWith[0].displayName.indexOf("Organization") > -1))
    //                 {
    //                     isGuestOrEveryone = true;
    //                     sharingType = sharedResult.SharingUserType;
    //                 }
    //             });

    //             // remove duplicates from the array (because of search results filling up the first positions of the array which is missing some data)
    //             sharedWithUsers = this.uniqForObject(sharedWithUsers);
                
    //             // if for some reason the sharing link data together with the search results produce one and the same user, return the first result including the UPN of the user in order for the livepersona to work
    //             if (sharedWithUsers.length == 1)
    //             {
    //                 _sharedWithSpecified.push(files[0]);
    //                 return;
    //             }
    //             else
    //             {
    //                 // fetch all the rest of the sharing link details using the last object in the array (because of search results filling up the first positions of the array which is missing some data)
    //                 const shareLink = files[files.length - 1];
    //                 shareLink.SharedWith = sharedWithUsers;
    //                 shareLink.SharingUserType = isGuestOrEveryone ? sharingType : "Member";
    //                 _sharedWithSpecified.push(shareLink);
    //             }
    //         }
    //     });

    //     return _sharedWithSpecified;
    // }
    
    public rightTrim(sourceString:string,searchString:string) 
    { 
        for(;;) 
        {
            var pos = sourceString.lastIndexOf(searchString); 
            if(pos === sourceString.length -1)
            {
                var result  = sourceString.slice(0,pos);
                sourceString = result; 
            }
            else 
            {
                break;
            }
        } 
        return sourceString;  
    }
    public convertToGraphUserFromLinkKind(linkKind:number) : microsoftgraph.User
    {
        let _user:microsoftgraph.User = {};
        switch(linkKind)
        {
            case 2:_user.displayName = "Organization View";break;
            case 3:_user.displayName = "Organization Edit";break;
            case 4:_user.displayName = "Anonymous View";break;
            case 5:_user.displayName = "Anonymous Edit";break;
            default:break;
        }
        _user.userType = "Link";
        return _user;
    }

    public convertUserToFacePilePersona(user:IdentitySet) : IFacepilePersona
    {
        if (user['siteUser'] != null)
        {
            let siteUser = user['siteUser'];
            const _user: IFacepilePersona = 
            {
                data : (siteUser.loginName.indexOf('#ext') != -1) ? "Guest" : "Member",
                personaName : siteUser.displayName,
                name : siteUser.loginName.replace("i:0#.f|membership|", "")
            };
            return _user;
        }
        else if(user['siteGroup'] != null)
        {
            let siteGroup = user['siteGroup'];
            const _user: IFacepilePersona = 
            {
                data : "Group",
                personaName : siteGroup.displayName,
                name : siteGroup.loginName.replace("c:0t.c|tenant|", "")
            };
            return _user;
        }
        else
        {
            const _user: IFacepilePersona = 
            {
                name : user.user.id,
                data : (user.user.id == null) ? "Guest" : "Member",
                personaName: user.user.displayName
            };
            return _user;
        }
       
    }

    public convertToFacePilePersona(users:IdentitySet[]) : IFacepilePersona[]
    {
        const _users:IFacepilePersona[] = [];
        if (users.length > 1)
        {
            users.forEach((user) => 
            {  
                if (user['siteUser'] != null)
                {
                    let siteUser = user['siteUser'];
                    const _user: IFacepilePersona = 
                    {
                        data : (siteUser.loginName.indexOf('#ext') != -1) ? "Guest" : "Member",
                        personaName : siteUser.displayName,
                        name : siteUser.loginName.replace("i:0#.f|membership|", "")
                    };
                    _users.push(_user);
                }
                else
                {
                    const _user: IFacepilePersona = 
                    {
                        name : user.user.id,
                        data : (user.user.id == null) ? "Guest" : "Member",
                        personaName: user.user.displayName
                    };
                    _users.push(_user);
                }
            });
        }
        else
        {
            let user = users[0];
            if (user['siteUser'] != null)
            {
                let siteUser = user['siteUser'];
                const _user: IFacepilePersona = 
                {
                    data : (siteUser.loginName.indexOf('#ext') != -1) ? "Guest" : "Member",
                    personaName : siteUser.displayName,
                    name : siteUser.loginName.replace("i:0#.f|membership|", "")
                };
                _users.push(_user);
            }
            else
            {
                const _user: IFacepilePersona = 
                {
                    id : user.user.id,
                    data : (user.user.id == null) ? "Guest" : "Member",
                    personaName : user.user.displayName
                };
                _users.push(_user);
            }
        }

        return _users;
    }

    public GetSortingMenuItems(column: IColumn, onSortColumn: (column: IColumn, isSortedDescending: boolean) => void): IContextualMenuItem[] {
        let menuItems = [];
        if (column.data == Number) {
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
        else if (column.data == Date) {
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
    
    // private _IsValuePresented(currentValues: IContextualMenuItem[], newValue: string): boolean {

    //     for (let i = 0; i < currentValues.length; i++) {
    //         if (currentValues[i].key == newValue) {
    //             return true;
    //         }
    //     }
    //     return false;
    // }
/// this is used to process the SharedWithUsersOWSUSER output to get the userPrincipalName and userType 
public processUsers(users: string) :  IFacepilePersona[]
{
    const _users:microsoftgraph.User[] = [];
    
    if (users == null || users == undefined)
        return _users;

    if (users.match("\n\n"))
    {
        var allUsers = users.split("\n\n");
        allUsers.forEach(element => {
        const user: IFacepilePersona = {
            personaName: element.split("|")[1].trim(),
            data: (element.indexOf("#ext#") > -1) ? "Guest" : "Member",
            id: element.split("|")[0].trim()
        };
        _users.push(user)
        });
    }
    else
    {
        const user: IFacepilePersona = {
            personaName: users.split("|")[1].trim(),
            data: (users.indexOf("#ext#") > -1) ? "Guest" : "Member",
            id: users.split("|")[0].trim()
        };
        _users.push(user)
    }
    return _users;
}
}