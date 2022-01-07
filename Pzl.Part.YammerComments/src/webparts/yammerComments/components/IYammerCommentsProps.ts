import { IYammerService } from "../services/YammerService";

export interface IYammerCommentsProps {
  yammerService: IYammerService;
  communityId: string;
}
