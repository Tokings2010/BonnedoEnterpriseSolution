import * as React from 'react';
import {
    Label,
    NormalPeoplePicker,
    IPersonaProps,
    IBasePickerSuggestionsProps,
    ValidationState,
    Spinner,
    SpinnerSize,
} from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISPUser {
    id?: number | string;
    Id?: number;
    Title?: string;
    Email?: string;
    LoginName?: string;
}

export interface IPeoplePickerProps {
    titleText?: string;
    selectedUsers?: ISPUser[];
    onChange?: (users: ISPUser[]) => void;
    personSelectionLimit?: number;
    principalTypes?: number[];
    resolveDelay?: number;
    required?: boolean;
    disabled?: boolean;
    spHttpClient?: SPHttpClient;
    pageContext?: PageContext;
    context?: PageContext;
    webPartContext?: WebPartContext;
    /** Optional: Custom user source list to pull users from */
    userListName?: string;
    /** Optional: Filter to apply when fetching from custom list */
    userFilter?: string;
}

const suggestionProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: 'Suggested People',
    mostRecentlyUsedHeaderText: 'Suggested Contacts',
    noResultsFoundText: 'No results found. Try a different search term.',
    loadingText: 'Searching...',
    showRemoveButtons: true,
    suggestionsContainerAriaLabel: 'Suggested contacts',
};

const PeoplePicker: React.FC<IPeoplePickerProps> = ({
    titleText = 'Select People',
    selectedUsers = [],
    onChange,
    personSelectionLimit = 1,
    resolveDelay = 200,
    required = false,
    disabled = false,
    spHttpClient,
    pageContext,
    context,
    webPartContext,
    userListName,
    userFilter,
}) => {
    // Use context prop if provided, otherwise use pageContext
    const effectiveContext = context || pageContext;

    const [isLoading, setIsLoading] = React.useState(false);
    const [errorMessage, setErrorMessage] = React.useState<string | null>(null);

    const [selectedItems, setSelectedItems] = React.useState<IPersonaProps[]>(
        selectedUsers.map(user => ({
            id: (user.id || user.Id || 0).toString(),
            primaryText: user.Title || '',
            secondaryText: user.Email || '',
            imageInitials: (user.Title || '').split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase(),
        }))
    );

    // Update selected items when props change
    React.useEffect(() => {
        setSelectedItems(selectedUsers.map(user => ({
            id: (user.id || user.Id || 0).toString(),
            primaryText: user.Title || '',
            secondaryText: user.Email || '',
            imageInitials: (user.Title || '').split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase(),
        })));
    }, [selectedUsers]);

    // Helper function to escape OData filter strings
    const escapeODataString = (str: string): string => {
        return str.replace(/'/g, "''");
    };

    const onFilterChanged = React.useCallback(
        async (filterText: string, currentSelected: IPersonaProps[]): Promise<IPersonaProps[]> => {
            if (filterText.length < 2) {
                return [];
            }

            // Clear any previous error
            setErrorMessage(null);

            // If no spHttpClient provided, return empty
            if (!spHttpClient || !effectiveContext) {
                console.warn('PeoplePicker: Missing spHttpClient or context');
                return [];
            }

            setIsLoading(true);
            let users: IPersonaProps[] = [];
            const searchTerm = filterText.trim();
            const escapedSearchTerm = escapeODataString(searchTerm);

            try {
                const siteUrl = effectiveContext.web.absoluteUrl;

                // ========== METHOD 1: Use SharePoint User Information List (Primary) ==========
                // This returns SharePoint user IDs which can be used with SharePoint REST API
                try {
                    console.log('PeoplePicker: Trying User Information List');
                    let userInfoUrl = `${siteUrl}/_api/web/siteusers?$top=50`;

                    if (userFilter) {
                        userInfoUrl += `&$filter=${userFilter}`;
                    } else {
                        // Use startswith for case-insensitive search in SharePoint
                        userInfoUrl += `&$filter=startswith(Title,'${escapedSearchTerm}') or startswith(Email,'${escapedSearchTerm}')`;
                    }

                    const userInfoResponse = await spHttpClient.get(
                        userInfoUrl,
                        SPHttpClient.configurations.v1
                    );

                    if (userInfoResponse.ok) {
                        const userInfoData = await userInfoResponse.json();

                        if (userInfoData.value && userInfoData.value.length > 0) {
                            console.log('PeoplePicker: User Info List returned', userInfoData.value.length, 'users');
                            users = userInfoData.value
                                .filter((user: any) => user && user.Title)
                                .map((user: any) => ({
                                    id: user.Id?.toString() || user.Id || '',
                                    primaryText: user.Title || '',
                                    secondaryText: user.Email || '',
                                    imageInitials: (user.Title || '').split(' ').map((n: string) => n[0] || '').join('').substring(0, 2).toUpperCase(),
                                    optionalText: user.Email || '',
                                }));
                        }
                    }
                } catch (userInfoError) {
                    console.error('PeoplePicker: User Information List search failed:', userInfoError);
                }

                // ========== METHOD 2: Use SharePoint Search API with people source ==========
                // Fallback - searches entire SharePoint
                if (users.length === 0) {
                    try {
                        console.log('PeoplePicker: Trying SharePoint Search API');
                        const searchQuery = encodeURIComponent(`${searchTerm}*`);
                        const searchUrl = `${siteUrl}/_api/search/query?` +
                            `querytext='${searchQuery}'&` +
                            `rowlimit=30&` +
                            `sourceid='b09a7990-05ea-4af9-81e8-86b7c5c5f8d1'&` +
                            `trimduplicates=false`;

                        const searchResponse = await spHttpClient.get(
                            searchUrl,
                            SPHttpClient.configurations.v1
                        );

                        if (searchResponse.ok) {
                            const searchData = await searchResponse.json();

                            if (searchData.d?.query?.PrimaryQueryResult?.RelevantResults) {
                                const results = searchData.d.query.PrimaryQueryResult.RelevantResults;

                                if (results.Table?.Rows?.results) {
                                    const rows = results.Table.Rows.results;

                                    users = rows
                                        .filter((row: any) => {
                                            const cells = row.Cells?.results ? row.Cells.results : [];
                                            const email = cells.find((c: any) => c.Key === 'WorkEmail');
                                            return email && email.Value && email.Value.length > 0;
                                        })
                                        .map((row: any) => {
                                            const cells = row.Cells.results;
                                            const title = cells.find((c: any) => c.Key === 'Title')?.Value || '';
                                            const email = cells.find((c: any) => c.Key === 'WorkEmail')?.Value || '';
                                            const accountName = cells.find((c: any) => c.Key === 'AccountName')?.Value || '';
                                            const jobTitle = cells.find((c: any) => c.Key === 'JobTitle')?.Value || '';

                                            return {
                                                id: accountName || email,
                                                primaryText: title,
                                                secondaryText: jobTitle ? `${email} - ${jobTitle}` : email,
                                                imageInitials: title.split(' ').map((n: string) => n[0] || '').join('').substring(0, 2).toUpperCase(),
                                                optionalText: email,
                                            };
                                        });
                                }
                            }
                        }
                    } catch (searchError) {
                        console.error('PeoplePicker: Search API failed:', searchError);
                    }
                }

                // ========== METHOD 3: Use Microsoft Graph API (Last Resort) ==========
                // Only use if SharePoint methods fail - this returns Azure AD IDs not SharePoint IDs
                if (users.length === 0 && webPartContext?.msGraphClientFactory) {
                    try {
                        console.log('PeoplePicker: Attempting Microsoft Graph API search for:', searchTerm);
                        const graphClient = await webPartContext.msGraphClientFactory.getClient('3');

                        // Use filter instead of $search - this works without ConsistencyLevel header
                        const graphResponse = await graphClient
                            .api('/users')
                            .filter(`startsWith(displayName,'${searchTerm}') or startsWith(mail,'${searchTerm}') or startsWith(givenName,'${searchTerm}') or startsWith(surname,'${searchTerm}')`)
                            .top(30)
                            .select('id,displayName,mail,jobTitle,givenName,surname,userPrincipalName')
                            .get();

                        if (graphResponse && graphResponse.value && graphResponse.value.length > 0) {
                            console.log('PeoplePicker: Graph API returned', graphResponse.value.length, 'users');
                            users = graphResponse.value
                                .filter((user: any) => user && user.displayName)
                                .map((user: any) => ({
                                    id: user.id || '',
                                    primaryText: user.displayName || '',
                                    secondaryText: user.jobTitle ? `${user.mail || user.userPrincipalName} - ${user.jobTitle}` : (user.mail || user.userPrincipalName || ''),
                                    imageInitials: (user.displayName || '').split(' ').map((n: string) => n[0] || '').join('').substring(0, 2).toUpperCase(),
                                    optionalText: user.mail || user.userPrincipalName || '',
                                }));
                        }
                    } catch (graphError) {
                        console.error('PeoplePicker: Graph API search failed:', graphError);
                    }
                }

                // ========== METHOD 4: Use custom user list if specified ==========
                if (users.length === 0 && userListName) {
                    try {
                        console.log('PeoplePicker: Trying custom user list:', userListName);
                        const listUrl = `${siteUrl}/_api/web/lists/getByTitle('${userListName}')/items?$top=30&$filter=contains(Title,'${escapedSearchTerm}')`;

                        const listResponse = await spHttpClient.get(
                            listUrl,
                            SPHttpClient.configurations.v1
                        );

                        if (listResponse.ok) {
                            const listData = await listResponse.json();

                            if (listData.value && listData.value.length > 0) {
                                users = listData.value
                                    .filter((item: any) => item.Title || item.Name)
                                    .map((item: any) => ({
                                        id: item.ID?.toString() || item.Id?.toString() || '',
                                        primaryText: item.Title || item.Name || '',
                                        secondaryText: item.Email || item.EmailAddress || '',
                                        imageInitials: (item.Title || item.Name || '').split(' ').map((n: string) => n[0] || '').join('').substring(0, 2).toUpperCase(),
                                        optionalText: item.Email || item.EmailAddress || '',
                                    }));
                            }
                        }
                    } catch (listError) {
                        console.error('PeoplePicker: Custom list search failed:', listError);
                    }
                }

                // If no users found after all methods, set error message
                if (users.length === 0) {
                    console.log('PeoplePicker: No users found with any method');
                    setErrorMessage('No users found. Please check that you have proper permissions to search users.');
                }

                // Filter out already selected users
                const selectedIds = new Set(currentSelected.map(item => item.id));
                const filteredUsers = users.filter(user => !selectedIds.has(user.id));

                // Remove duplicates based on id
                const uniqueUsers = filteredUsers.reduce((unique: IPersonaProps[], item) => {
                    const exists = unique.find(u => u.id === item.id);
                    if (!exists) {
                        unique.push(item);
                    }
                    return unique;
                }, []);

                return uniqueUsers;

            } catch (error) {
                console.error('PeoplePicker search error:', error);
                setErrorMessage('Error searching for users. Please check console for details.');
                return [];
            } finally {
                setIsLoading(false);
            }
        },
        [spHttpClient, effectiveContext, webPartContext, userListName, userFilter]
    );

    const onItemsChanged = React.useCallback(
        (items: IPersonaProps[]): void => {
            // Limit selection
            if (items.length > personSelectionLimit) {
                items = items.slice(0, personSelectionLimit);
            }

            setSelectedItems(items);

            // Convert to ISPUser format and notify parent
            if (onChange) {
                const spUsers: ISPUser[] = items.map(item => {
                    // Try to parse as number, but keep original if not possible
                    const numericId = parseInt(item.id || '0', 10);
                    return {
                        // Use numeric ID if parseable, otherwise keep the original string (for Graph/SharePoint IDs)
                        id: isNaN(numericId) ? item.id : numericId,
                        Title: item.primaryText || '',
                        Email: item.secondaryText || '',
                        LoginName: item.id || '',
                    };
                });
                onChange(spUsers);
            }
        },
        [personSelectionLimit, onChange]
    );

    const getTextFromItem = (item: IPersonaProps): string => {
        return item.primaryText || '';
    };

    const validateInput = (input: string): ValidationState => {
        if (input.length > 0 && selectedItems.length >= personSelectionLimit) {
            return ValidationState.invalid;
        }
        return ValidationState.valid;
    };

    return (
        <div>
            {titleText && (
                <Label required={required}>
                    {titleText}
                </Label>
            )}
            {isLoading && (
                <div style={{ marginBottom: '4px' }}>
                    <Spinner size={SpinnerSize.small} label="Searching..." />
                </div>
            )}
            {errorMessage && !isLoading && (
                <div style={{
                    marginBottom: '4px',
                    padding: '8px',
                    backgroundColor: '#fff4f4',
                    border: '1px solid #e0a0a0',
                    borderRadius: '4px',
                    color: '#c42b1c',
                    fontSize: '12px'
                }}>
                    {errorMessage}
                </div>
            )}
            <NormalPeoplePicker
                onResolveSuggestions={onFilterChanged}
                getTextFromItem={getTextFromItem}
                onChange={onItemsChanged}
                selectedItems={selectedItems}
                pickerSuggestionsProps={{
                    ...suggestionProps,
                    loadingText: isLoading ? 'Searching...' : suggestionProps.loadingText,
                }}
                inputProps={{
                    placeholder: required ? `Select up to ${personSelectionLimit} person(s)` : 'Type to search for people...',
                }}
                itemLimit={personSelectionLimit}
                disabled={disabled}
                onValidateInput={validateInput}
                resolveDelay={resolveDelay}
                selectionAriaLabel="Selected contacts"
                styles={{
                    root: {
                        width: '100%',
                    },
                }}
            />
        </div>
    );
};

export default PeoplePicker;
