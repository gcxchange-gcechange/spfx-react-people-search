import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from "./ReactDirectory.module.scss";
import { PersonaCard } from "./PersonaCard/PersonaCard";
import { spservices } from "../../../SPServices/spservices";
import { IReactDirectoryState } from "./IReactDirectoryState";
import { SelectLanguage } from "./SelectLanguage";
import {
    Spinner, SpinnerSize, MessageBar, MessageBarType, SearchBox, Icon, Label,
    Pivot, PivotItem, PivotLinkFormat, PivotLinkSize, Dropdown, IDropdownOption, IStackItemStyles, Image, IImageProps, ImageFit, PrimaryButton, IStyleSet, IPivotStyles
} from "office-ui-fabric-react";
import { Stack, IStackStyles, IStackTokens } from 'office-ui-fabric-react/lib/Stack';

import { debounce } from "throttle-debounce";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { ISPServices } from "../../../SPServices/ISPServices";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { spMockServices } from "../../../SPServices/spMockServices";
import { IReactDirectoryProps } from './IReactDirectoryProps';
import Paging from './Pagination/Paging';
import ReactHtmlParser from 'react-html-parser';
import * as _ from 'lodash';

const slice: any = require('lodash/slice');
const filter: any = require('lodash/filter');
const wrapStackTokens: IStackTokens = { childrenGap: 30 };



const DirectoryHook: React.FC<IReactDirectoryProps> = (props) => {
    const strings = SelectLanguage(props.prefLang);
    let _services: ISPServices = null;
    if (Environment.type === EnvironmentType.Local) {
        _services = new spservices(props.context);
    } else {
        _services = new spservices(props.context);
    }
    const [az, setaz] = useState<string[]>([]);
    const [alphaKey, setalphaKey] = useState<string>('A');
    const [state, setstate] = useState<IReactDirectoryState>({
        users: [],
        isLoading: true,
        errorMessage: "",
        hasError: false,
        indexSelectedKey: 'A',
        searchString: "FirstName",
        searchText: ""
    });

    const color = props.context.microsoftTeams ? "white" : "";
    const hidingUsers: string[] =props.hidingUsers && props.hidingUsers.length > 0 ? props.hidingUsers.split(",") : [];
    console.log("hidingUsers",hidingUsers);
    // Paging
    const [pagedItems, setPagedItems] = useState<any[]>([]);
    const [pageSize, setPageSize] = useState<number>(props.pageSize ? props.pageSize : 10);
    const [currentPage, setCurrentPage] = useState<number>(1);

    const _onPageUpdate = async (pageno?: number) => {
        var currentPge = (pageno) ? pageno : currentPage;
        var startItem = ((currentPge - 1) * pageSize);
        var endItem = currentPge * pageSize;
        let filItems = slice(state.users, startItem, endItem);
        setCurrentPage(currentPge);
        setPagedItems(filItems);
    };

    const diretoryGrid =
        pagedItems && pagedItems.length > 0
            ? pagedItems.map((user: any) => {
                return (
                    <PersonaCard
                        context={props.context}
                        prefLang={props.prefLang}
                        profileProperties={{
                            DisplayName: user.PreferredName,
                            Title: user.JobTitle,
                            PictureUrl: user.PictureURL,
                            Email: user.WorkEmail,
                            Department: user.Department,
                            WorkPhone: user.WorkPhone,
                            Location: user.OfficeNumber
                                ? user.OfficeNumber
                                : user.BaseOfficeLocation
                        }}
                    />
                );
            })
            : [];
    const _loadAlphabets = () => {
        let alphabets: string[] = [];
        for (let i = 65; i < 91; i++) {
            alphabets.push(
                String.fromCharCode(i)
            );
        }
        setaz(alphabets);
    };

    const _alphabetChange = async (item?: PivotItem, ev?: React.MouseEvent<HTMLElement>) => {
        setstate({ ...state, searchText: "", indexSelectedKey: item.props.itemKey, isLoading: true });
        setalphaKey(item.props.itemKey);
        setCurrentPage(1);
    };
    const _searchByAlphabets = async (initialSearch: boolean) => {
        setstate({ ...state, isLoading: true, searchText: '' });
        let users = null;
        if (initialSearch) {
            if (props.searchFirstName)
                users = await _services.searchUsersNew(
                  props.context,
                  "",
                  `FirstName:a*`,
                  false,
                  hidingUsers
                );
            else users = await _services.searchUsersNew(
              props.context,
              "a",
              "",
              true,
              hidingUsers
            );
        } else {
            if (props.searchFirstName)
                users = await _services.searchUsersNew(
                  props.context,
                  "",
                  `FirstName:${alphaKey}*`,
                  false,
                  hidingUsers
                );
            else users = await _services.searchUsersNew(
              props.context,
              `${alphaKey}`,
              "",
              true,
              hidingUsers
            );
        }
        setstate({
            ...state,
            searchText: '',
            indexSelectedKey: initialSearch ? 'A' : state.indexSelectedKey,
            users:
                users && users.PrimarySearchResults
                    ? users.PrimarySearchResults
                    : null,
            isLoading: false,
            errorMessage: "",
            hasError: false
        });
    };

    let _searchUsers = async () => {

        try {
            setstate({
                ...state,
                isLoading: true,

            });
            const searchText = state.searchText;
            if (searchText.length > 0) {
                let searchProps: string[] = props.searchProps && props.searchProps.length > 0 ?
                    props.searchProps.split(',') : ['FirstName', 'LastName', 'PreferredName',];

                let qryText: string = '';
                let finalSearchText: string = searchText ? searchText.replace(/ /g, '+') : searchText;
                if (props.clearTextSearchProps) {
                    let tmpCTProps: string[] = props.clearTextSearchProps.indexOf(',') >= 0 ? props.clearTextSearchProps.split(',') : [props.clearTextSearchProps];
                    if (tmpCTProps.length > 0) {
                        searchProps.map((srchprop, index) => {
                            let ctPresent: any[] = filter(tmpCTProps, (o) => { return o.toLowerCase() == srchprop.toLowerCase(); });
                            if (ctPresent.length > 0) {
                                if (index == searchProps.length - 1) {
                                    qryText += `${srchprop}:${searchText}*`;
                                } else qryText += `${srchprop}:${searchText}* OR `;
                            } else {
                                if (index == searchProps.length - 1) {
                                    qryText += `${srchprop}:${finalSearchText}*`;
                                } else qryText += `${srchprop}:${finalSearchText}* OR `;
                            }
                        });
                    } else {
                        searchProps.map((srchprop, index) => {
                            if (index == searchProps.length - 1)
                                qryText += `${srchprop}:${finalSearchText}*`;
                            else qryText += `${srchprop}:${finalSearchText}* OR `;
                        });
                    }
                } else {
                    searchProps.map((srchprop, index) => {
                        if (index == searchProps.length - 1)
                            qryText += `${srchprop}:${finalSearchText}*`;
                        else qryText += `${srchprop}:${finalSearchText}* OR `;
                    });
                }
                console.log(qryText);
                const users = await _services.searchUsersNew(
                  props.context,
                  "",
                  qryText,
                  false,
                  hidingUsers
                );
                setstate({
                    ...state,
                    searchText: searchText,
                    indexSelectedKey: null,
                    users:
                        users && users.PrimarySearchResults
                            ? users.PrimarySearchResults
                            : null,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });
            } else {
                setstate({ ...state, searchText: '' });
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    const _searchBoxChanged = (newvalue: string): void => {
        setCurrentPage(1);
        setstate({
            ...state,
            searchText: newvalue
        }

        );
    };

     _searchUsers = _.debounce(_searchUsers,500,);



    useEffect(() => {
        setPageSize(props.pageSize);
        if (state.users) _onPageUpdate();
    }, [state.users, props.pageSize]);

    useEffect(() => {
        if (alphaKey.length > 0 && alphaKey != "0") _searchByAlphabets(false);
    }, [alphaKey]);

    useEffect(() => {
        _loadAlphabets();
        _searchByAlphabets(true);
    }, [props]);

    const itemAlignmentsStackTokens: IStackTokens = {
        childrenGap: 20,
    };
    const stackItemStyles: IStackItemStyles = {
        root: {
            paddingTop: 5,
        },

    };

    const imageProps: Partial<IImageProps> = {
        imageFit: ImageFit.centerContain,
        width: 200,
        height: 200,
        src: require("../assets/HidingYeti.png")
    };
    const piviotStyles: Partial<IStyleSet<IPivotStyles>>={
        link:{
            backgroundColor: "#ccc",      
        }   
            
    };

    return (
        <div className={styles.reactDirectory}>
            <WebPartTitle displayMode={props.displayMode} title={props.title}
                updateProperty={props.updateProperty} />
            <div className={styles.searchBox}>
                <Stack horizontal tokens={itemAlignmentsStackTokens}>
                    <Stack.Item order={1} styles={stackItemStyles}>
                        <span><label>{strings.SearchBoxLabel}</label></span>
                    </Stack.Item>
                    <Stack.Item order={2} >
                        <SearchBox placeholder={strings.SearchPlaceHolder} className={styles.searchTextBox}
                            onSearch={_searchUsers}
                            value={state.searchText}
                            onChange={_searchBoxChanged} />
                    </Stack.Item>
                    <Stack.Item order={2} >
                        <PrimaryButton onClick={_searchUsers}>{strings.SearchButtonLabel}</PrimaryButton>
                    </Stack.Item>
                </Stack>

                <div>
                    {<Pivot styles={piviotStyles} className={styles.alphabets} linkFormat={PivotLinkFormat.tabs}
                        selectedKey={state.indexSelectedKey} onLinkClick={_alphabetChange}
                        linkSize={PivotLinkSize.normal} >
                        {az.map((index: string) => {
                            return (
                                <PivotItem headerText={index} itemKey={index} key={index} />
                            );
                        })}
                    </Pivot>}
                </div>
            </div>
            {state.isLoading ? (
                <div style={{ marginTop: '10px' }}>
                    <Spinner size={SpinnerSize.large} label={strings.LoadingText} />
                </div>
            ) : (
                <>
                    {state.hasError ? (
                        <div style={{ marginTop: '10px' }}>
                            <MessageBar messageBarType={MessageBarType.error}>
                                {state.errorMessage}
                            </MessageBar>
                        </div>
                    ) : (
                        <>

                            {!pagedItems || pagedItems.length == 0 ? (
                                <div className={styles.noUsers}>

                                    <Stack horizontal tokens={itemAlignmentsStackTokens}>
                                        <Stack.Item order={1} styles={stackItemStyles} >
                                            <span>
                                                <Image {...imageProps} alt={strings.NoUserFoundImageAltText} />
                                            </span>
                                        </Stack.Item>
                                        <Stack.Item order={2} >
                                            <span>
                                                <p>{ReactHtmlParser(strings.DirectoryMessage)}</p>
                                                <PrimaryButton href={strings.NoUserFoundEmail}>{strings.NoUserFoundLabelText}</PrimaryButton>
                                            </span>
                                        </Stack.Item>
                                    </Stack>
                                </div>
                            ) : (
                                <>
                                    <div style={{ width: '100%', display: 'inline-block' }}>
                                        <Paging
                                            totalItems={state.users.length}
                                            itemsCountPerPage={pageSize}
                                            onPageUpdate={_onPageUpdate}
                                            currentPage={currentPage} />
                                    </div>

                                    <Stack horizontal horizontalAlign="center" wrap tokens={wrapStackTokens}>
                                        <div>{diretoryGrid}</div>
                                    </Stack>
                                    <div style={{ width: '100%', display: 'inline-block' }}>
                                        <Paging
                                            totalItems={state.users.length}
                                            itemsCountPerPage={pageSize}
                                            onPageUpdate={_onPageUpdate}
                                            currentPage={currentPage} />
                                    </div>
                                </>
                            )}
                        </>
                    )}
                </>
            )}
        </div>
    );
};

export default DirectoryHook;
