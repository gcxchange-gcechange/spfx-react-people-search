import { PeoplePickerEntity } from '@pnp/sp';

export interface ISPServices {
  searchUsers(searchString: string, searchFirstName: boolean);
  searchUsersNew(
    context: any,
    searchString: string,
    srchQry: string,
    isInitialSearch: boolean,
    hidingUsers: any,
    pageNumber?: number
  );
}