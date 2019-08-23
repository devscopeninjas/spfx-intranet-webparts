import { IGraphGroup } from "./IGraphGroup";

export interface IGraphConsumerState {
    groups: Array<IGraphGroup>;
    searchFor: string;
  }