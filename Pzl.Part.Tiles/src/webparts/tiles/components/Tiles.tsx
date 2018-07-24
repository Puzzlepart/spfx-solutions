import * as React from 'react';
import styles from './Tiles.module.scss';
import { ITilesProps } from './ITilesProps';
import * as pnp from "sp-pnp-js/lib/pnp";
import { Spinner, SpinnerType } from "office-ui-fabric-react/lib/Spinner";

export interface ITilesState {
  items?: Array<any>;
  isLoading?: boolean;
}

export default class Tiles extends React.Component<ITilesProps, ITilesState> {
  constructor(props) {
    super(props);
    this.state = {
      items: [],
      isLoading: true,
    };
  }
  public componentDidMount(): void {
    this.fetchData();
  }
  public componentDidUpdate(prevprops): void {
    if (this.props !== prevprops) {
      this.fetchData();
    }
  }
  public render(): React.ReactElement<ITilesProps> {
    let { isLoading, items }: ITilesState = this.state;
    let elements = items.map((item: any, index: number) => {
      return <a className={styles.promotedLink} style={{ width: `${this.props.imageWidth}px`, height: `${this.props.imageHeight}px` }} key={index} target={(item[this.props.newTabField]) ? "_blank" : ""} href={(item[this.props.linkField]) ? item[this.props.linkField].Url : "#"}>
        <img className={styles.image} src={(item[this.props.backgroundImageField]) ? item[this.props.backgroundImageField].Url : this.props.fallbackImageUrl} />
        <div className={styles.textArea} style={{ height: `${this.props.imageHeight}px`, top: `${this.props.imageHeight / 3 * 2}px` }}>
          <div className={styles.container} style={{ padding: `${this.props.textPadding}px` }}>
            <div className={styles.title}>{item.Title}</div>
            <div className={styles.description}>{item[this.props.descriptionField]}</div>
          </div>
        </div>
      </a>;
    });
    if (isLoading) {
      return <Spinner type={SpinnerType.large} />;
    } else {
      return <div>{
        elements.length > 0 && <div className={styles.promotedLinks}>
          {elements}
        </div>}
      </div>;

    }
  }
  private async fetchData(): Promise<void> {
    try {
      let filter = (this.props.tileTypeField && this.props.tileType) ? `${this.props.tileTypeField} eq '${this.props.tileType}'` : '';
      let response = await pnp.sp.web.lists.getByTitle(this.props.list).items.filter(filter).orderBy((this.props.orderByField) ? this.props.orderByField : "ID").top(this.props.count).get();
      this.setState({
        items: response,
        isLoading: false,
      });
    } catch (error) {
      this.setState({
        isLoading: false,
      });
      throw error;
    }
  }
}
