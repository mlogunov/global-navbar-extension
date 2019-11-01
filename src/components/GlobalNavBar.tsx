import * as React from 'react';
import { IGlobalNavBarProps } from './IGlobalNavBarProps';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { IContextualMenuItem, ContextualMenuItemType } from 'office-ui-fabric-react/lib/ContextualMenu';
import { ISPTermObject } from '../services/SPTermStoreService';

export const GlobalNavBar: React.StatelessComponent<IGlobalNavBarProps> = (props: IGlobalNavBarProps): React.ReactElement<IGlobalNavBarProps> => {

    const menuItem = (item: ISPTermObject, itemType: ContextualMenuItemType): IContextualMenuItem => {
        return ({
            key: item.identity,
            name: item.name,
            itemType: itemType,
            href: item.localCustomProperties['_Sys_Nav_SimpleLinkUrl'],
            subMenuProps: item.terms.length > 0 ? 
                {items: item.terms.map((i: ISPTermObject): IContextualMenuItem => {return menuItem(i, ContextualMenuItemType.Normal)})}
                : null
        })
    }

    const commandBatItems: IContextualMenuItem[] = props.menuItems.map((item: ISPTermObject): IContextualMenuItem => {
        return menuItem(item, ContextualMenuItemType.Header)
    })

    return <CommandBar items={commandBatItems} />
}