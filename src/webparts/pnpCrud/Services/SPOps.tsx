import * as React from 'react';
import { Item, sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { People } from '../components/PnpCrud';






 export  class SPOperations{

     async GetAllItems(): Promise<People[]> {
         let x: People[]=[];
         const items: any[] = await sp.web.lists.getByTitle('Employee OnBoard').items.getAll().then((item: any) => {
             let e = new People(item.Id, item.DepartmentId, item.Jobtype, item.ManagerId, item.Lastname, item.Role,item.Title);
             console.log(item);
             
             x.push(e);
             return items;
         });
         
         
        
         return x;
     }


 }