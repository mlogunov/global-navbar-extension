import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { sp } from "@pnp/sp";
import { ITermStore, taxonomy, ITermSet, ITerm, ITermData } from "@pnp/sp-taxonomy";

export interface ISPTermObject {
    identity: string;
    isAvailableForTagging: boolean;
    name: string;
    guid: string;
    customSortOrder: string;
    terms: ISPTermObject[];
    localCustomProperties: any;
}

export class SPTermStoreService {
    constructor(private context: ApplicationCustomizerContext){
        sp.setup({
            spfxContext: this.context
        })
    }

    public async getGlobalNavItemsAsync(termSet: string): Promise<ISPTermObject[]> {
        let items: ISPTermObject[] = [];
        try{
            const stores: ITermStore[] = await taxonomy.termStores.get();
            const set: ITermSet = await stores[0].getTermSetsByName(termSet, 1033).getByName(termSet);
            const terms: (ITermData & ITerm)[] = await set.terms.get();
            const rootTerms: (ITermData & ITerm)[] = terms.filter((term: ITermData) => term.IsRoot);
            if(rootTerms && rootTerms.length > 0){
                items = await Promise.all<ISPTermObject>(
                    rootTerms.map(async (term: (ITermData & ITerm)): Promise<ISPTermObject> => {
                        return await this.getNavItemAsync(term)
                    })
                ) 
            }
        }
        catch(error){
            console.log(error);
            return Promise.reject(error);
        }
        
        return items;
    }

    private async getNavItemAsync(term: (ITermData & ITerm)): Promise<ISPTermObject> {
        return(
            {
                identity: term.Id,
                isAvailableForTagging: term.IsAvailableForTagging,
                name: term.Name,
                guid: term.Id,
                customSortOrder: term.CustomSortOrder,
                terms: await this.getChildNavItemsAsync(term),
                localCustomProperties: term.LocalCustomProperties
            }
        )
    }

    private async getChildNavItemsAsync(term: ITerm): Promise<ISPTermObject[]> {
        let items: ISPTermObject[] = [];
        const terms: ITerm[] = await term.terms.get();
        if(terms && terms.length > 0){
            items = await Promise.all<ISPTermObject>(
                terms.map(async (term: ITerm): Promise<ISPTermObject> => {
                    return await this.getNavItemAsync(term)
                })
            ) 
        }
        return items;
    }
}