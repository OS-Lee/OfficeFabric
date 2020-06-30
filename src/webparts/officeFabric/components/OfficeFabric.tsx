import * as React from 'react';
import styles from './OfficeFabric.module.scss';
import { IOfficeFabricProps } from './IOfficeFabricProps';
import { escape } from '@microsoft/sp-lodash-subset';

// import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
//import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TextField, MaskedTextField } from 'office-ui-fabric-react/lib/TextField';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { ChoiceGroup } from 'office-ui-fabric-react/lib/ChoiceGroup';
import {
  ComboBox, Fabric, IComboBoxOption, mergeStyles,
  SelectableOptionMenuItemType, Toggle
} from 'office-ui-fabric-react/lib/index';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { DefaultButton, PrimaryButton, Callout, Icon, DirectionalHint, Link, getTheme, FontWeights, mergeStyleSets, getId } from 'office-ui-fabric-react';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IListItem {
  Id: number;
  Title: string;
  Body: {
    Body_p1: string;
    Body_p2: string;
  };
}

export interface IReactExState {
  controlName: string,
  ControlValue: string,
  windowsID: number,
  showPanel: boolean,
  items: IListItem[],
  Options: IComboBoxOption[],
  selectionDetails: {},
  showModal: boolean,
  autoComplete: boolean,
  allowFreeform: boolean,
  isCalloutVisible?: boolean;
  calloutTitle: string;
  calloutContent: string;
  calloutLink: string;
}
const theme = getTheme();
const CustomStyles = mergeStyleSets({
  IconArea: {
    verticalAlign: 'top',
    display: 'inline-block',
    textAlign: 'center',
    margin: '0 0px'
  },
  callout: {  
    minWidth:'260px',
    display:"inline"
  },
  header: {
    //padding: '18px 24px 12px',
    height: '30px',
    backgroundColor: '#b3b3cc'    
  },
  footer: {
    //padding: '18px 24px 12px',
    height: '50px',
    backgroundColor: '#b3b3cc'
  },
  title: [
    theme.fonts.xLarge,
    {
      margin: 0,
      color: theme.palette.neutralPrimary,
      fontWeight: FontWeights.semilight
    }
  ],
  footerTitle:{
    fontWeight: 'bold'
  },
  inner: {
    height: '100%',
    padding: '0 24px 20px'
  },
  actions: {
    position: 'relative',
    marginTop: 20,
    width: '100%',
    whiteSpace: 'nowrap'
  },
  subtext: [
    theme.fonts.small,
    {
      margin: 0,
      color: theme.palette.neutralPrimary,
      fontWeight: FontWeights.semilight
    }
  ],
  link: [
    theme.fonts.medium,
    {
      color: theme.palette.neutralPrimary
    }
  ]
});

export default class OfficeFabric extends React.Component<IOfficeFabricProps, IReactExState> {

  private _selection: Selection;
  private _allItems: IListItem[];
  private _columns: IColumn[];

  public constructor(props: IOfficeFabricProps, state: IReactExState) {
    super(props);
    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() })
    });
    this._allItems = [
      {
        Id: 0,
        Title: "test0",
        Body: {
          Body_p1: "test0_p1",
          Body_p2: "test0_p2"
        },
      },
      {
        Id: 1,
        Title: "test1",
        Body: {
          Body_p1: "test1_p1",
          Body_p2: "test1_p2"
        }
      }
    ];
    this._columns = [
      { key: 'Id', name: 'Id', fieldName: 'Id', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'Title', name: 'Title', fieldName: 'Title', minWidth: 200, maxWidth: 400, isResizable: true },
      {
        key: 'Body', name: 'Body', minWidth: 200, maxWidth: 400, isResizable: true, onRender: (item) => (
          <div>
            {item.Body.Body_p1}
          </div>)
      },
    ];
    this.state = {
      controlName: "",
      ControlValue: "",
      windowsID: 123.4,
      showPanel: false,
      items: this._allItems,
      Options: [],
      selectionDetails: this._getSelectionDetails(),
      showModal: false,
      autoComplete: false,
      allowFreeform: true,
      isCalloutVisible: false,
      calloutTitle: "",
      calloutContent: "",
      calloutLink: ""
    };

  }

  private _menuButtonElement: HTMLElement | null;
  private _labelId: string = getId('callout-label');
  private _descriptionId: string = getId('callout-description');

  private _showModal = (): void => {
    this.setState({ showModal: true });
  };

  private _closeModal = (): void => {
    this.setState({ showModal: false });
  };

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IListItem).Title;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onItemInvoked = (item: IListItem): void => {
    alert(`Item invoked: ${item.Title}`);
  }

  private _showPanel = () => {
    this.setState({ showPanel: true });
  }

  private _hidePanel = () => {
    this.setState({ showPanel: false });
  }
  private _Save = () => {
    //to do save logic
    this.setState({ showPanel: false });
    alert('save clicked');
  }
  private _onRenderFooterContent = () => {
    return (
      <div>
        <PrimaryButton onClick={this._hidePanel} style={{ marginRight: '8px' }}>
          Save
        </PrimaryButton>
        <DefaultButton onClick={this._showPanel}>Cancel</DefaultButton>
      </div>
    );
  };

  public componentDidMount() {
    var reactHandler = this;
    this.props.context.spHttpClient.get(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('TestList')/items?select=ID,Title`,
      SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
        response.json().then((responseJSON: any) => {
          let tempOptions: IComboBoxOption[] = [];
          responseJSON.value.forEach(element => {
            tempOptions.push({ key: element.ID, text: element.Title })
          });
          reactHandler.setState({
            Options: tempOptions
          });
        });
      });
  }
  // private _onJobTitReportToChange = (ev: React.FormEvent<HTMLInputElement>, newValue?: string) => {
  //   this.props.onJobTitleReportToChange(newValue);
  // }

  public handleObjectWithMultipleFields = (ev, newText: string): void => {
    const target = ev.target;
    const value = newText;
    var _ControlName = target.name;

    this.setState({
      controlName: _ControlName,
      ControlValue: value
    });
  }
  private _onShowMenuClicked = (title: any): void => {
    this._menuButtonElement = document.getElementById(title);
    if (title == "icon1") {
      this.setState({
        isCalloutVisible: !this.state.isCalloutVisible,
        calloutTitle: "My CallOut 1",
        calloutContent: "<div><p>Description</p><div>My CallOut Content 1</div></div>"+
        "<div><p>Valid Values</p><div>free text ...blabla</div></div>",
        calloutLink: "<a href='http://bing.com'>Go to Bing1</a>",
      });
    }
    if (title == "icon2") {
      this.setState({
        isCalloutVisible: !this.state.isCalloutVisible,
        calloutTitle: "My CallOut 2",
        calloutContent: "My CallOut Content 2",
        calloutLink: "<a href='http://microsoft.com'>Go to microsoft</a>"
      });
    }
    if (title == "icon3") {
      this.setState({
        isCalloutVisible: !this.state.isCalloutVisible,
        calloutTitle: "My CallOut 3",
        calloutContent: "My CallOut Content 3",
        calloutLink: "<a href='http://microsoft.com'>Go to microsoft</a>"
      });
    }
    if (title == "icon4") {
      this.setState({
        isCalloutVisible: !this.state.isCalloutVisible,
        calloutTitle: "My CallOut 4",
        calloutContent: "My CallOut Content 4",
        calloutLink: "<a href='http://microsoft.com'>Go to microsoft</a>"
      });
    }
  };

  private _onCalloutDismiss = (): void => {
    this.setState({
      isCalloutVisible: false
    });
  };
  public render(): React.ReactElement<IOfficeFabricProps> {
    return (
      <div className={styles.officeFabric}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>

              <div>
                <div id="icon1" className={CustomStyles.IconArea} ref={menuButton => (this._menuButtonElement = menuButton)}>
                  <Icon iconName='Info' onClick={e => this._onShowMenuClicked(e.currentTarget.title)} title="icon1" />
                </div>

                <div id="icon2" style={{ marginLeft: "400px" }} className={CustomStyles.IconArea} ref={menuButton => (this._menuButtonElement = menuButton)}>
                  <Icon iconName='Info' onClick={e => this._onShowMenuClicked(e.currentTarget.title)} title="icon2" />
                </div>
                <br /> <br /> <br /> <br /> <br /> <br /> <br /> <br /> <br /> <br /> <br /> <br /> <br /> <br /> <br />
                <div id="icon3" className={CustomStyles.IconArea} ref={menuButton => (this._menuButtonElement = menuButton)}>
                  <Icon iconName='Info' onClick={e => this._onShowMenuClicked(e.currentTarget.title)} title="icon3" />
                </div>
                <div id="icon4" style={{ marginLeft: "400px" }} className={CustomStyles.IconArea} ref={menuButton => (this._menuButtonElement = menuButton)}>
                  <Icon iconName='Info' onClick={e => this._onShowMenuClicked(e.currentTarget.title)} title="icon4" />
                </div>
                {this.state.isCalloutVisible && (
                  <Callout
                    className={CustomStyles.callout}
                    ariaLabelledBy={this._labelId}
                    ariaDescribedBy={this._descriptionId}
                    role="alertdialog"
                    gapSpace={0}
                    directionalHint={DirectionalHint.rightCenter}
                    target={this._menuButtonElement}
                    onDismiss={this._onCalloutDismiss}
                    setInitialFocus={true}
                    style={{display:'inline'}}
                  >
                    <div className={CustomStyles.header}>                      
                        {this.state.calloutTitle}                      
                    </div>
                    <div className={CustomStyles.inner}>
                      <div dangerouslySetInnerHTML={{ __html: this.state.calloutContent}}>                      
                      </div>
                      <div className={CustomStyles.actions}>
                        <div dangerouslySetInnerHTML={{ __html: this.state.calloutLink }} />
                      </div>
                    </div>
                    <div className={CustomStyles.footer}>
                      <p className={CustomStyles.footerTitle} id={this._labelId}>
                        Owner:content
                      </p>                      
                      <div>
                      <Icon iconName='mail' title="Email" />Request Change:xxx
                      </div>
                    </div>
                  </Callout>
                )}
              </div>

              <ComboBox
                label="ComboBox with toggleable freeform/auto-complete"
                key={'' + this.state.autoComplete + this.state.allowFreeform /*key causes re- 
            render when toggles change*/}
                allowFreeform={this.state.allowFreeform}
                autoComplete={this.state.autoComplete ? 'on' : 'off'}
                options={this.state.Options}
              />
              <TextField name="txtA" value={this.state.windowsID.toString()}
                onChange={this.handleObjectWithMultipleFields} />

              <TextField name="txtB"
                onChange={this.handleObjectWithMultipleFields} />
              <div>
                <DefaultButton secondaryText="Opens the Sample Panel" onClick={this._showPanel} text="Open Panel" />
                <Panel
                  isOpen={this.state.showPanel}
                  type={PanelType.smallFixedFar}
                  onDismiss={this._hidePanel}
                  headerText="Panel - Small, right-aligned, fixed, with footer"
                  closeButtonAriaLabel="Close"
                  onRenderFooterContent={this._onRenderFooterContent}
                >
                  <DefaultButton className={styles.tablink} text="Button" />
                  <ChoiceGroup
                    options={[
                      {
                        key: 'A',
                        text: 'Option A'
                      },
                      {
                        key: 'B',
                        text: 'Option B',
                        checked: true
                      },
                      {
                        key: 'C',
                        text: 'Option C',
                        disabled: true
                      },
                      {
                        key: 'D',
                        text: 'Option D',
                        checked: true,
                        disabled: true
                      }
                    ]}
                    label="Pick one"
                    required={true}
                  />
                </Panel>

                <DefaultButton secondaryText="Opens the Sample Modal" onClick={this._showModal} text="Open Modal" />

              </div>
              <MarqueeSelection selection={this._selection}>
                <DetailsList
                  items={this.state.items}
                  columns={this._columns}
                  setKey="set"
                  layoutMode={DetailsListLayoutMode.justified}
                  selection={this._selection}
                  selectionPreservedOnEmptyClick={true}
                  ariaLabelForSelectionColumn="Toggle selection"
                  ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                  checkButtonAriaLabel="Row checkbox"
                  onItemInvoked={this._onItemInvoked}
                />
              </MarqueeSelection>
              <a href="https://aka.ms/spfx" className={styles.button}>
                <span className={styles.label}>Learn more</span>
              </a>
              <br/>
              Edit Icon
              <Icon iconName="Edit" className={styles.iconStyles} style={{ cursor: 'pointer'}} />
              <br/>
              <Icon iconName='Mail' />
              <br />
              <Icon iconName='CirclePlus' />
              <br />
              <Icon iconName='LocationDot' />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
