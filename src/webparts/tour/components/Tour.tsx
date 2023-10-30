import * as React from "react";
import styles from "./Tour.module.scss";
import Tours from "reactour";
import { CompoundButton, Icon, Link } from "office-ui-fabric-react";
import { TourHelper } from "./TourHelper";
import { disableBodyScroll, enableBodyScroll } from "body-scroll-lock";

export interface ITourProps {
  description: string;
  actionValue: string;
  icon: string;
  collectionData: any[];
  onClose: () => void;
  editMode: boolean;
  isFullWidth: boolean;
}

export interface ITourState {
  isTourOpen: boolean;
  steps: any[];
  tourDisabled: boolean;
}

export default class Tour extends React.Component<ITourProps, ITourState> {
  constructor(props: ITourProps) {
    super(props);
    this.state = {
      isTourOpen: false,
      steps: [],
      tourDisabled: true,
    };
  }

  public componentDidMount() {
    this.setState({
      steps: TourHelper.getTourSteps(this.props.collectionData),
    });
    if (
      this.props.collectionData != undefined &&
      this.props.collectionData.length > 0
    ) {
      this.setState({ tourDisabled: false });
    }
  }

  public componentDidUpdate(newProps) {
    if (
      JSON.stringify(this.props.collectionData) !=
      JSON.stringify(newProps.collectionData)
    ) {
      this.setState({
        steps: TourHelper.getTourSteps(this.props.collectionData),
      });
      if (
        this.props.collectionData != undefined &&
        this.props.collectionData.length > 0
      ) {
        this.setState({ tourDisabled: false });
      } else {
        this.setState({ tourDisabled: true });
      }
    }
  }

  public render(): React.ReactElement<ITourState> {
    const { description, actionValue, icon, editMode, onClose, isFullWidth } =
      this.props;

    const iconProps = icon ? { iconName: icon } : null;

    const tour = (
      <Tours
        onRequestClose={this._closeTour}
        startAt={0}
        steps={this.state.steps}
        isOpen={this.state.isTourOpen}
        maskClassName="mask"
        className={styles.reactTourCustomCss}
        accentColor={"#00A961"}
        rounded={5}
        onAfterOpen={this._disableBody}
        onBeforeClose={this._enableBody}
      />
    );

    if (isFullWidth) {
      return (
        <div className={styles.fullWidthTour}>
          {!editMode && (
            <Icon
              title="Close this webpart"
              iconName="RemoveFilter"
              className={styles.closeButton}
              onClick={onClose}
            />
          )}
          <div className={styles.tourLink}>
            {iconProps && <Icon {...iconProps} className={styles.icon} />}
            {this.state.tourDisabled && (
              <div className={styles.actionDisabled}>{actionValue}</div>
            )}
            {!this.state.tourDisabled && (
              <div onClick={this._openTour} className={styles.action}>
                {actionValue}
              </div>
            )}
            <div className={styles.description}>{description}</div>
          </div>
          {tour}
        </div>
      );
    } else {
      return (
        <div className={styles.tour}>
          <CompoundButton
            primary
            text={actionValue}
            secondaryText={description}
            disabled={this.state.tourDisabled}
            onClick={this._openTour}
            checked={this.state.isTourOpen}
            iconProps={iconProps}
            className={styles.tutorialButton}
          ></CompoundButton>
          {!editMode && (
            <Icon
              title="Close this webpart"
              iconName="RemoveFilter"
              className={styles.closeButton}
              onClick={onClose}
            />
          )}
          {tour}
        </div>
      );
    }
  }

  private _disableBody = (target) => disableBodyScroll(target);
  private _enableBody = (target) => enableBodyScroll(target);

  private _closeTour = () => {
    this.setState({ isTourOpen: false });
  }

  private _openTour = () => {
    if (!this.state.tourDisabled) {
      this.setState({ isTourOpen: true });
    }
  }
}
