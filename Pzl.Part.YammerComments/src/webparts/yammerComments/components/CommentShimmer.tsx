import * as React from "react";
import { IShimmerElement, Shimmer, ShimmerElementType } from "office-ui-fabric-react";

export interface ICommentShimmerProps {
    isReply: boolean;
}

export const CommentShimmer: React.FunctionComponent<ICommentShimmerProps> = (props) => {

    const user: IShimmerElement = { type: ShimmerElementType.circle, height: 32, verticalAlign: 'top' };
    const gap: IShimmerElement = { type: ShimmerElementType.gap, width: 16 };
    const text: IShimmerElement = { type: ShimmerElementType.line, height: 64, verticalAlign: 'top' };

    const elements: IShimmerElement[] = !props.isReply ? [user, gap, text] :
        [{ type: ShimmerElementType.gap, width: 48 }, user, gap, text];

    return (
        <Shimmer shimmerElements={elements} />
    );
};