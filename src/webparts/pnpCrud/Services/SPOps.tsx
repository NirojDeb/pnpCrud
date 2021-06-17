import * as React from 'react';
import { Item, sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import { People } from '../components/PnpCrud';






 export  class SPOperations{

     async GetAllItems(): Promise<People[]> {
         let x: People[]=[];
         const items: any[] = await sp.web.lists.getByTitle('Employee OnBoard').items.getAll();
         await items.forEach(async (item) => {
             
             let e = new People(item.Id, item.DepartmentId, item.Jobtype, item.ManagerId, item.Lastname, item.Role, item.Title);
             x.push(e);
         });
         return x;
     }

     async GetUser(id): Promise<any> {
         return await sp.web.siteUsers.getById(id).get();
     }

     async GetRoles():Promise<any> {
         return await sp.web.lists.getByTitle('Employee OnBoard').fields.getByInternalNameOrTitle('Role').select('Choices,ID').get();
     }

     async GetJobTypes(): Promise<any> {
         return await sp.web.lists.getByTitle('Employee OnBoard').fields.getByInternalNameOrTitle('Jobtype').select('Choices,ID').get();
     }

     async CreateItem(emp:People): Promise<any> {
         var newItem = {
             Title: emp.Title,
             Lastname:emp.Lastname,
             Jobtype: emp.Jobtype,
             DepartmentId: emp.DepartmentId,
             Role: emp.Role,
             ManagerId: emp.ManagerId
         }

         await sp.web.lists.getByTitle('Employee OnBoard').items.add(newItem);
     }

     async GetUserByLoginName(loginName: string): Promise<any> {
         return await sp.web.siteUsers.getByLoginName(loginName).get();
     }


     async DeleteItem(id: number): Promise<any> {
         return await sp.web.lists.getByTitle('Employee OnBoard').items.getById(id).delete();
     }


     async EditItem(id: number, emp: People): Promise<any> {
         var newItem = {
             Title: emp.Title,
             Lastname: emp.Lastname,
             Jobtype: emp.Jobtype,
             DepartmentId: emp.DepartmentId,
             Role: emp.Role,
             ManagerId: emp.ManagerId
         }

         return await sp.web.lists.getByTitle('Employee OnBoard').items.getById(id).update(newItem);
     }

     async GetDepartements(): Promise<any> {
         var res = await sp.web.lists.getByTitle('Department').items.select("Title", "Id").get();
         var mapping = [];
         res.forEach((r) => {
             let dep = [r.Title, r.Id];
             mapping.push(dep);
         });
         return mapping;
     }

  }