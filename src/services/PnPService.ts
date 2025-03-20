/* eslint-disable @typescript-eslint/no-explicit-any */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { getSP } from "./PnPSP";
import { SPFI } from "@pnp/sp";
import { IList } from "@pnp/sp/lists";
import { createBatch } from "@pnp/sp/batching";

export class PnPService {

    private context: WebPartContext;
    private sp: SPFI;

    public constructor(context: WebPartContext, sourceURL?: string) {
        this.context = context;
        if (sourceURL === undefined) {
            this.sp = getSP(this.context);
        }
        else {
            this.sp = getSP(this.context, sourceURL);
        }
    }

    public async getAllItemsByTitle(listTitle: string, select?: string[], orderby?: string, ascending?: boolean, expand?: string, filterQuery?: string): Promise<any[]> {
        let response: any[] = [];
        const pageSize = 4999;
        let page = 0;
        const selectFields: string = select ? select.join(`,`) : ``;
        const filter: string = filterQuery !== undefined ? filterQuery : ``;
        let result = await this.sp.web.lists.getByTitle(listTitle).items.
            select(selectFields).
            expand(expand !== undefined && expand !== `undefined` ? expand : ``).
            filter(filter).
            orderBy(orderby ? orderby : ``, ascending).top(pageSize)();
        response = response.concat(result);
        while (result.length === pageSize) {
            page++;
            result = await this.sp.web.lists.getByTitle(listTitle).items.
                select(selectFields).
                expand(expand !== undefined && expand !== `undefined` ? expand : ``).
                filter(filter).
                orderBy(orderby ? orderby : ``, ascending).top(pageSize).skip(page * pageSize)();
            response = response.concat(result);
        }
        return Promise.resolve(response.length > 0 ? response : []);
    }

    public async getItemsInBatch(listTitle: string, itemIds: number[]): Promise<any[]> {
        const result: any[] = [];
        const list: IList = this.sp.web.lists.getByTitle(listTitle);
        const [batchedListBehavior, execute] = createBatch(list);
        list.using(batchedListBehavior);

        const resItems = itemIds.map((id) => {
            return list.items.getById(id)().then(r => result.push(r));
        });
        await execute();
        await Promise.all(resItems);
        return Promise.resolve(result);
    }

    public async createItemsInBatch(listTitle: string, body: any[], noOfItemsToBeCreated: number): Promise<string> {
        try {
            const list: IList = this.sp.web.lists.getByTitle(listTitle);
            const [batchedListBehavior, execute] = createBatch(list);
            list.using(batchedListBehavior);
            body.map((item) => {
                list.items.add(item).catch((err) => {
                    console.log(err);
                });
            });
            await execute();
            return Promise.resolve(`Ok`);
        } catch (err) {
            console.log(err);
            return Promise.resolve(`Error`);
        }
    }
}
