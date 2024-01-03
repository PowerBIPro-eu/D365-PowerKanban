import * as React from "react";
import { Card, ICardTokens, ICardSectionStyles, ICardSectionTokens } from '@uifabric/react-cards';
import { Tile } from "./Tile";
import { BoardLane } from "../domain/BoardLane";
import { Metadata, Attribute } from "../domain/Metadata";
import { CardForm } from "../domain/CardForm";
import { useDrop } from "react-dnd";
import { ItemTypes } from "../domain/ItemTypes";
import { Option } from "../domain/Metadata";
import { BoardViewConfig, BoardEntity } from "../domain/BoardViewConfig";
import { Subscription } from "../domain/Subscription";

interface LaneProps {
    config: BoardEntity;
    cardForm: CardForm;
    dndType?: string;
    lane: BoardLane;
    metadata: Metadata;
    minWidth?: string;
    notifications?: {[key: string]: Array<Notification>};
    refresh: () => Promise<void>;
    searchText: string;
    selectedSecondaryForm?: CardForm;
    separatorMetadata: Attribute;
    subscriptions: {[key: string]: Array<Subscription>};
    isSecondaryLane?: boolean;
    openRecord: (reference: Xrm.LookupValue) => void;
    selectedRecords?: {[key: string]: boolean};
}

const LaneRender = (props: LaneProps) => {
    const [{ canDrop, isOver }, drop] = useDrop({
      accept: props.dndType ?? ItemTypes.Tile,
      drop: () => ({ option: props.lane.option }),
      collect: monitor => ({
        isOver: monitor.isOver(),
        canDrop: monitor.canDrop(),
      }),
      canDrop: (item, monitor) => {
        const typedItem = item as unknown as { id: string; sourceLane: Option } | undefined;

        if (!typedItem.sourceLane._parsedTransitionData) {
          return true;
        }

        return typedItem.sourceLane._parsedTransitionData.some(p => p.to === props.lane.option.Value);
      }
    });

    const borderColor = props.lane.option.Color ?? "#3b79b7";

    const isActive = canDrop && isOver;
    let style: React.CSSProperties = { };

    if (isActive) {
      style = { borderWidth: "3px", borderStyle: "dashed", borderColor: "#02f01c" };
    } else if (canDrop) {
      style = { borderWidth: "3px", borderStyle: "dashed", borderColor: "#3b79b7" };
    }

    const mapDataToTile = ((d: any) => <Tile
      dndType={props.dndType}
      laneOption={props.lane.option}
      borderColor={borderColor}
      metadata={props.metadata}
      cardForm={props.cardForm}
      key={`tile_${d[props.metadata.PrimaryIdAttribute]}`}
      refresh={props.refresh}
      subscriptions={props.subscriptions[d[props.metadata.PrimaryIdAttribute]] ?? []}
      searchText={props.searchText}
      data={d}
      config={props.config}
      separatorMetadata={props.separatorMetadata}
      openRecord={props.openRecord}
      isSecondary={props.isSecondaryLane}
      isSelected={props.selectedRecords && props.selectedRecords[d[props.metadata.PrimaryIdAttribute]]} />
    );

    return (
        <div ref={drop} style={{ ...style, minWidth: !props.config.fitLanesToScreenWidth ? (props.minWidth ?? "400px") : (props.lane.data.length ? "auto" : "20px"), marginTop: "5px", marginBottom: "5px", marginLeft: "2.5px", marginRight: "2.5px", flex: "1" }}>
            <Card tokens={{ childrenGap: "5px" }} styles={{ root: { maxWidth: "auto", minWidth: "auto", backgroundColor: "#fff", borderColor: "#d8d8d8", height: "100%", borderTopStyle: "solid", borderTopColor: borderColor, borderTopWidth: "3px", color: "#333333" }}}>
                <Card.Section styles={{ root: { padding: "5px", borderBottom: "1px solid rgba(0,0,0,.125)" }}}>
                  <h5 title={props.lane.option.Label.UserLocalizedLabel.Label} style={{ fontSize: "17.5px", fontWeight: "normal", cursor: "default", color: "#045999", textOverflow: "ellipsis", overflow: "hidden", whiteSpace: "nowrap"}}>{props.lane.option.Label.UserLocalizedLabel.Label} { !props.config.hideCountOnLane && <a>({props.lane.data.length})</a> }</h5>
                </Card.Section>
                <Card.Section styles={ { root: { overflow: "auto", padding: "0.5rem" } } }>
                    {
                      props.cardForm &&
                      props.lane.data.map(mapDataToTile)
                    }
                </Card.Section>
            </Card>
        </div>
    );
};

export const Lane = React.memo(LaneRender, (a, b) => {
  if (a.cardForm != b.cardForm) {
      return false;
  }

  if (a.dndType != b.dndType) {
      return false;
  }

  if (a.metadata != b.metadata) {
      return false;
  }

  if (a.searchText != b.searchText) {
      return false;
  }

  if (a.notifications != b.notifications) {
      return false;
  }

  if (a.subscriptions != b.subscriptions) {
      return false;
  }

  if (a.lane.data.length != b.lane.data.length) {
      return false;
  }

  if (a.selectedRecords != b.selectedRecords) {
    return false;
  }

  return true;
});