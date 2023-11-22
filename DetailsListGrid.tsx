import * as React from 'react';
import { Announced } from '@fluentui/react/lib/Announced';
import { TextField, ITextFieldStyles } from '@fluentui/react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from '@fluentui/react/lib/DetailsList';
import { MarqueeSelection } from '@fluentui/react/lib/MarqueeSelection';
import { mergeStyles } from '@fluentui/react/lib/Styling';
import { Text } from '@fluentui/react/lib/Text';
import { mergeStyleSets, DefaultButton, FocusTrapZone, Layer, Overlay, Popup } from '@fluentui/react';
import { ActionButton } from '@fluentui/react/lib/Button';
import { IIconProps } from '@fluentui/react';
import { Label } from '@fluentui/react/lib/Label';
import { IconButton } from '@fluentui/react/lib/Button';
import { IInputs } from './generated/ManifestTypes';

const exampleChildClass = mergeStyles({
    display: 'block',
    marginBottom: '10px',
});

const textFieldStyles: Partial<ITextFieldStyles> = { root: { maxWidth: '300px' } };
const addIcon: IIconProps = { iconName: 'Add' };

/**
  * CSS Style Start
  */
// const popupStyles = mergeStyleSets({
//     root: {
//         background: 'rgba(0, 0, 0, 0.2)',
//         bottom: '0',
//         left: '0',
//         position: 'fixed',
//         right: '0',
//         top: '0',
//     },
//     content: {
//         background: 'white',
//         left: '50%',
//         maxWidth: '400px',
//         padding: '0 2em 2em',
//         position: 'absolute',
//         top: '50%',
//         transform: 'translate(-50%, -50%)',
//     },
// });

// const popUpButnStyle = mergeStyles({
//     display: 'flex',
//     justifyContent: 'center',
// });

// const okBtnStyle = mergeStyles({
//     marginRight: '10px', // Adjust the margin-right to control the space between buttons
// });

// const addResourceBtnStyle = mergeStyles({
//     display: 'left',
//     justifyContent: 'left-start', // This aligns the content to the left
//     border: 'none',
// });

// const noRowsLabelClass = mergeStyles({
//     fontFamily: '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
//     fontWeight: '400',
//     color: 'rgb(50, 49, 48)',
// });

/**
  * CSS Style End
  */

export interface IDetailsResourceAvailableItem {
    key: number;
    resourceName: string;
    startTime: string;
    endTime: string;
    resourceId: string;
    resourceType: string
}

export interface IDetailsListResourceAvailableProps {
    context: ComponentFramework.Context<IInputs>;
}

export interface IDetailsListResourceAvailableState {
    items: IDetailsResourceAvailableItem[];
    selectionDetails: string;
    disabled?: boolean;
    isPopupVisible?: boolean;
    currentPage: number;
    itemsPerPage: number;
}

interface IParameters {
    SelectedResources: string;
    OperationType: string;
}

export class DetailsListResourceAvailable extends React.Component<IDetailsListResourceAvailableProps, IDetailsListResourceAvailableState> {
    private _selection: Selection;
    private _allItems: IDetailsResourceAvailableItem[] = [];
    private _columns: IColumn[];

    constructor(props: IDetailsListResourceAvailableProps) {
        super(props);
        const context = this.props.context;
        this._selection = new Selection({
            onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
        });

        this.initializeData();

        this._columns = [
            { key: 'resourceName', name: 'Resource Name', fieldName: 'resourceName', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'startTime', name: 'Start Time', fieldName: 'startTime', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'endTime', name: 'End Time', fieldName: 'endTime', minWidth: 100, maxWidth: 200, isResizable: true },
            { key: 'resourceType', name: 'Resource Type', fieldName: 'resourceType', minWidth: 100, maxWidth: 200, isResizable: true }
        ];
    }

    componentWillMount() {
        this.state = {
            itemsPerPage: 4,
            items: this._allItems.length == 0 ? this._allItems : this._allItems.slice(0, this.state.itemsPerPage),
            selectionDetails: this._getSelectionDetails(),
            disabled: true,
            isPopupVisible: false,
            currentPage: 1,
        };
    }
    public render(): React.JSX.Element {
        const { items, selectionDetails, itemsPerPage, currentPage } = this.state;
        const { disabled } = this.state;
        const totalPages = Math.ceil(this._allItems.length / itemsPerPage);


        return (
            <div>
                {/* <ActionButton iconProps={addIcon} allowDisabledFocus disabled={disabled} onClick={this._showPopup} className={addResourceBtnStyle}>Add Resource</ActionButton> */}
                {/* <DefaultButton iconProps={addIcon} allowDisabledFocus disabled={disabled} onClick={this._showPopup} className={addResourceBtnStyle}>Add Resource</DefaultButton> */}
                {/* <div className={exampleChildClass}>{selectionDetails}</div> */}
                {/* <Announced message={selectionDetails} /> */}
                {/* <TextField
                    className={exampleChildClass}
                    label="Filter by name:"
                    onChange={this._onFilter}
                    styles={textFieldStyles}
                /> */}
                {/* <Announced message={`Number of items after filter applied: ${items.length}.`} /> */}
                <MarqueeSelection selection={this._selection}>
                    <div>{this._allItems.length == 0 || this._allItems.length < 0 ? (
                        <>
                            <DefaultButton iconProps={addIcon} allowDisabledFocus disabled={true} onClick={this._showPopup} className="addResourceBtnStyle">Add Resource</DefaultButton>
                            <Label className="noRowsLabelClass">No data available</Label>
                        </>
                    ) : (
                        <>
                            <DefaultButton iconProps={addIcon} allowDisabledFocus disabled={disabled} onClick={this._showPopup} className="addResourceBtnStyle">Add Resource</DefaultButton>
                            <DetailsList
                                items={items}
                                columns={this._columns}
                                setKey="set"
                                layoutMode={DetailsListLayoutMode.justified}
                                selection={this._selection}
                                selectionPreservedOnEmptyClick={true}
                                ariaLabelForSelectionColumn="Toggle selection"
                                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                                checkButtonAriaLabel="select row"
                                onItemInvoked={this._onItemInvoked} />
                            {totalPages > 1 && (
                                <div className="pagination">
                                    <IconButton
                                        iconProps={{ iconName: 'Back' }}
                                        title="Previous"
                                        ariaLabel="Previous"
                                        onClick={this._goToPreviousPage}
                                        disabled={currentPage === 1}
                                    />
                                    <Text variant={'small'}>Page {currentPage} of {totalPages}</Text>
                                    <IconButton
                                        iconProps={{ iconName: 'Forward' }}
                                        title="Forward"
                                        ariaLabel="Forward"
                                        onClick={this._goToNextPage}
                                        disabled={currentPage === totalPages}
                                    />
                                </div>
                            )}
                        </>
                    )
                    }
                    </div>
                </MarqueeSelection>

                {this.state.isPopupVisible && (
                    <Layer>
                        <Popup
                            className="root"
                            role="dialog"
                            aria-modal="true"
                            onDismiss={this._hidePopup}
                        >
                            <Overlay onClick={this._hidePopup} />
                            <FocusTrapZone>
                                <div role="document" className="content">
                                    <h2>Confirm Resource Booking</h2>
                                    <p>
                                        Do you want to confirm resource booking ?
                                    </p>
                                    <div className="popUpButnStyle">
                                        <DefaultButton onClick={() => this._addResourceClick(this.props.context)} className="okBtnStyle">OK</DefaultButton>
                                        <DefaultButton onClick={this._hidePopup}>Cancel</DefaultButton>
                                    </div>
                                </div>
                            </FocusTrapZone>
                        </Popup>
                    </Layer>
                )}
            </div>
        );
    }

    private _getSelectionDetails(): string {
        const selectionCount = this._selection.getSelectedCount();

        switch (selectionCount) {
            case 0:
                this.setState({
                    disabled: true,
                });
                return 'No items selected';
            case 1:
                this.setState({
                    disabled: false,
                });
                return '1 item selected: ' + (this._selection.getSelection()[0] as IDetailsResourceAvailableItem).resourceName;
            default:
                this.setState({
                    disabled: false,
                });
                return `${selectionCount} items selected`;
        }
    }

    private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
        this.setState({
            items: text ? this._allItems.filter(i => i.resourceName.toLowerCase().indexOf(text) > -1) : this._allItems,
        });
    };

    private _onItemInvoked = (item: IDetailsResourceAvailableItem): void => {
        alert(`Item invoked: ${item.resourceName}`);
    };

    private _addResourceClick = (context : any) => {
        const selectedItems = this._selection.getSelection() as IDetailsResourceAvailableItem[];
        const selectedKeys = selectedItems.map(item => item.key);
        const unselectedItems = this._allItems.filter(item => !selectedKeys.includes(item.key));
        this._allItems = unselectedItems;
        const totalPages = Math.ceil(this._allItems.length / this.state.itemsPerPage);

        this.setState({
            items: this._allItems.length == 0 ? this._allItems : this._allItems.slice(0, this.state.itemsPerPage),
            isPopupVisible: false,
            currentPage: 1,
        });

        selectedItems.forEach(selectedItem => {
            let selectedItemsString = JSON.stringify(selectedItem);
            this._createBookableResourceBooking(selectedItemsString, context);
        });
    }

    private _showPopup = () => {
        this.setState({ isPopupVisible: true });
    }

    private _hidePopup = () => {
        this.setState({ isPopupVisible: false });
    }

    private _goToPreviousPage = () => {
        if (this.state.currentPage > 1) {
            this.setState({
                currentPage: this.state.currentPage - 1,
                items: this._allItems.length == 0 ? this._allItems : this._allItems.slice(
                    (this.state.currentPage - 2) * this.state.itemsPerPage,
                    (this.state.currentPage - 1) * this.state.itemsPerPage
                ),
            });
        }
    }

    private _goToNextPage = () => {
        const totalPages = Math.ceil(this._allItems.length / this.state.itemsPerPage);
        if (this.state.currentPage < totalPages) {
            this.setState({
                currentPage: this.state.currentPage + 1,
                items: this._allItems.length == 0 ? this._allItems : this._allItems.slice(
                    this.state.currentPage * this.state.itemsPerPage,
                    (this.state.currentPage + 1) * this.state.itemsPerPage
                ),
            });
        }
    }

    private async initializeData() {
        try {
            let resourceAvailable: any;
            let listOfAvailableResource;

            resourceAvailable = await this._getTimeSlotDetails(this.props.context);

            if (resourceAvailable != null && resourceAvailable != undefined) {
                listOfAvailableResource = JSON.parse(resourceAvailable)

                if (listOfAvailableResource != null && listOfAvailableResource != undefined &&
                    listOfAvailableResource.length > 0) {
                    for (let i = 0; i < listOfAvailableResource.length; i++) {
                        this._allItems.push({
                            key: i,
                            resourceName: listOfAvailableResource[i]["ResourceName"],
                            startTime: listOfAvailableResource[i]["StartTime"],
                            endTime: listOfAvailableResource[i]["EndTime"],
                            resourceId: listOfAvailableResource[i]["ResourceId"],
                            resourceType: listOfAvailableResource[i]["ResourceType"]
                        });
                    }
                }
            }
            this.setState({
                items: this._allItems.length === 0 ? this._allItems : this._allItems.slice(0, this.state.itemsPerPage),
            });

        } catch (error) {
            this.setState({
                items: []
            });
        }
    }

    private _createBookableResourceBooking = (selectedItemsString: string, context: any) => {
        const clientUrl = context.page.getClientUrl();
        const formId = context.page.entityId;

        // Parameters
        var parameters: IParameters = {
            SelectedResources: selectedItemsString,
            OperationType: "BookableResourceBooking"
        }

        var req = new XMLHttpRequest();
        req.open("POST", `${clientUrl}/api/data/v9.2/msdyn_workorders(` + formId + `)/Microsoft.Dynamics.CRM.acc_FitDispatchResourceAvailabilty`, true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("Accept", "application/json");
        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                req.onreadystatechange = null;
                if (this.status === 200 || this.status === 204) {
                    var result = JSON.parse(this.response);
                    // Return Type: mscrm.acc_FitDispatchResourceAvailabiltyResponse
                    // Output Parameters
                    var resourceavailabilty = result["ResourceAvailabilty"]; // Edm.String
                } else {
                    console.log(this.responseText);
                }
            }
        };
        req.send(JSON.stringify(parameters));
    }

    private _getTimeSlotDetails = (context: any) => {
        return new Promise(function (resolve, reject) {
            // Parameters
            var parameters: IParameters = {
                SelectedResources: "",
                OperationType: "RetriveTimeSlot"
            }// Edm.String

            const clientUrl = context.page.getClientUrl();
            const formId = context.page.entityId;
            let resourceavailabilty = null
            var req = new XMLHttpRequest();
            req.open("POST", `${clientUrl}/api/data/v9.2/msdyn_workorders(` + formId + `)/Microsoft.Dynamics.CRM.acc_FitDispatchResourceAvailabilty`, false);
            req.setRequestHeader("OData-MaxVersion", "4.0");
            req.setRequestHeader("OData-Version", "4.0");
            req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
            req.setRequestHeader("Accept", "application/json");
            req.onreadystatechange = function () {
                if (this.readyState === 4) {
                    req.onreadystatechange = null;
                    if (this.status === 200 || this.status === 204) {
                        var result = JSON.parse(this.response);
                        // Return Type: mscrm.acc_FitDispatchResourceAvailabiltyResponse
                        // Output Parameters
                        resourceavailabilty = result["ResourceAvailabilty"]; // Edm.String
                        resolve(resourceavailabilty);
                    } else {
                        resolve(null);
                    }
                }
            };
            req.send(JSON.stringify(parameters));
        });
    }
}
