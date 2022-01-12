import IUser from "./IUser";

export default interface IComment {    
    id?: string;
    text: string;
    groupId: string;
    replyToId?: string;
    user?: IUser;
    created?: Date;
    replies?: IComment[];
}