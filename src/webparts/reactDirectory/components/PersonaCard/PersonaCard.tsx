import * as React from "react";
import styles from "./PersonaCard.module.scss";
import { IPersonaCardProps } from "./IPersonaCardProps";
import { IPersonaCardState } from "./IPersonaCardState";
import { Log, Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { SPComponentLoader } from "@microsoft/sp-loader";

import {
  Persona,
  PersonaSize,
  DocumentCard,
  DocumentCardType,
  Icon,
  HoverCard,
  IPlainCardProps,
  DefaultButton,
  HoverCardType,
  Stack,
  IStackTokens,
  IStackStyles,
  DefaultPalette,
  IStackItemStyles,
} from "office-ui-fabric-react";

const EXP_SOURCE: string = "SPFxDirectory";
const LIVE_PERSONA_COMPONENT_ID: string =
  "914330ee-2df2-4f6e-a858-30c23a812408";

export class PersonaCard extends React.Component<
  IPersonaCardProps,
  IPersonaCardState
> {
  constructor(props: IPersonaCardProps) {
    super(props);

    this.state = { livePersonaCard: undefined, pictureUrl: undefined };
  }
  /**
   *
   *
   * @memberof PersonaCard
   */
  public async componentDidMount() {
    if (Environment.type !== EnvironmentType.Local) {
      const sharedLibrary = await this._loadSPComponentById(
        LIVE_PERSONA_COMPONENT_ID
      );
      const livePersonaCard: any = sharedLibrary.LivePersonaCard;
      this.setState({ livePersonaCard: livePersonaCard });
    }
  }

  /**
   *
   *
   * @param {IPersonaCardProps} prevProps
   * @param {IPersonaCardState} prevState
   * @memberof PersonaCard
   */
  public componentDidUpdate(
    prevProps: IPersonaCardProps,
    prevState: IPersonaCardState
  ): void {}

  /**
   *
   *
   * @private
   * @returns
   * @memberof PersonaCard
   */
  private _LivePersonaCard() {
    return React.createElement(
      this.state.livePersonaCard,
      {
        serviceScope: this.props.context.serviceScope,
        upn: this.props.profileProperties.Email,
        onCardOpen: () => {
          console.log("LivePersonaCard Open");
        },
        onCardClose: () => {
          console.log("LivePersonaCard Close");
        },
      },
      this._PersonaCard()
    );
  }

  /**
   *
   *
   * @private
   * @returns {JSX.Element}
   * @memberof PersonaCard
   */
  private _PersonaCard(): JSX.Element {
    return (
      <DocumentCard
        className={styles.documentCard}
        type={DocumentCardType.normal}
      >
        <div className={styles.persona}>
          <Persona
            text={this.props.profileProperties.DisplayName}
            secondaryText={this.props.profileProperties.Title}
            tertiaryText={this.props.profileProperties.Department}
            imageUrl={this.props.profileProperties.PictureUrl}
            size={PersonaSize.size72}
            imageShouldFadeIn={false}
            imageShouldStartVisible={true}
          >
            {this.props.profileProperties.Email ? (
              <div className={styles.textOverflow}>
                <Icon
                  iconName="PublicEmail"
                  style={{ fontSize: "12px", verticalAlign: "sub" }}
                />
                <span style={{ marginLeft: 5, fontSize: "12px" }}>
                  {" "}
                  {this.props.profileProperties.Email}
                </span>
              </div>
            ) : (
              ""
            )}
            {this.props.profileProperties.WorkPhone ? (
              <div>
                <Icon iconName="Phone" style={{ fontSize: "12px" }} />
                <span style={{ marginLeft: 5, fontSize: "12px" }}>
                  {" "}
                  {this.props.profileProperties.WorkPhone}
                </span>
              </div>
            ) : (
              ""
            )}
            {this.props.profileProperties.Location ? (
              <div className={styles.textOverflow}>
                <Icon iconName="Poi" style={{ fontSize: "12px" }} />
                <span style={{ marginLeft: 5, fontSize: "12px" }}>
                  {" "}
                  {this.props.profileProperties.Location}
                </span>
              </div>
            ) : (
              ""
            )}
          </Persona>
        </div>
      </DocumentCard>
    );
  }
  /**
   * Load SPFx component by id, SPComponentLoader is used to load the SPFx components
   * @param componentId - componentId, guid of the component library
   */
  private async _loadSPComponentById(componentId: string): Promise<any> {
    try {
      const component: any = await SPComponentLoader.loadComponentById(
        componentId
      );
      return component;
    } catch (error) {
      Promise.reject(error);
      Log.error(EXP_SOURCE, error, this.props.context.serviceScope);
    }
  }

  private onRenderPlainCard = (): JSX.Element => {
    // Tokens definition
    const containerStackTokens: IStackTokens = { childrenGap: 5 };
    const horizontalGapStackTokens: IStackTokens = {
      childrenGap: 20,
    };
    const itemAlignmentsStackTokens: IStackTokens = {
      childrenGap: 20,
    };
    // Styles definition
    const stackStyles: IStackStyles = {
      root: {
        padding: 0,
      },
    };
    const stackItemStyles: IStackItemStyles = {
      root: {
        padding: 5,
      },
    };
    return (
      <div className={styles.customPopup}>
        <Stack>
          <Stack horizontal tokens={itemAlignmentsStackTokens}>
            <Stack.Item align="auto" styles={stackItemStyles}>
              <div className={styles.customPopupImg}>
                <img
                  src={this.props.profileProperties.PictureUrl}
                  alt={this.props.profileProperties.DisplayName}
                />
              </div>
            </Stack.Item>
            <Stack.Item align="stretch" styles={stackItemStyles}>
              <div>
                <p className={styles.customPopupDisplayName}>
                  {this.props.profileProperties.DisplayName}
                </p>
              </div>
            </Stack.Item>
          </Stack>
          <Stack horizontal tokens={itemAlignmentsStackTokens}>
            <Stack.Item align="center" styles={stackItemStyles}>
              <div>
                <a
                  href={`mailto:${this.props.profileProperties.Email}`}
                  target="_blank"
                >
                  <Icon
                    iconName="PublicEmail"
                    style={{ verticalAlign: "sub", marginRight: 5 }}
                  />
                  Send email
                </a>
              </div>
            </Stack.Item>
            <Stack.Item align="center" styles={stackItemStyles}>
              <div>
                <a
                  href={`MSTeams:/l/chat/0/0?users=${this.props.profileProperties.Email}`}
                >
                  <Icon
                    iconName="Chat"
                    style={{ verticalAlign: "sub", marginRight: 5 }}
                  />
                  Start chat
                </a>
              </div>
            </Stack.Item>
          </Stack>
        </Stack>
      </div>
    );
  };

  /**
   *
   *
   * @returns {React.ReactElement<IPersonaCardProps>}
   * @memberof PersonaCard
   */
  public render(): React.ReactElement<IPersonaCardProps> {
    const plainCardProps: IPlainCardProps = {
      onRenderPlainCard: this.onRenderPlainCard,
    };
    return (
      <div className={styles.personaContainer}>
        {
          <HoverCard
            instantOpenOnClick={true}
            cardDismissDelay={2000}
            type={HoverCardType.plain}
            plainCardProps={plainCardProps}
          >
            {this._PersonaCard()}
          </HoverCard>
        }
      </div>
    );
  }
}
