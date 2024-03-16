import * as WebApiClient from "xrm-webapi-client";

const avatarCache: { [key: string]: Promise<string> } = {};

export const FetchUserAvatar = (id: string) => {
    if (avatarCache[id]) {
        return avatarCache[id];
    }

    avatarCache[id] = WebApiClient.Retrieve({
        entityName: "systemuser",
        entityId: id,
        queryParams: "?$select=entityimage_url",
    }).then((data: any) => data.entityimage_url as string);

    console.log("avatar Cache", avatarCache);
    return avatarCache[id];
};
