import * as React from 'react';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import styles from './PnpCrud.module.scss';
import { IPnpCrudProps } from './IPnpCrudProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Log } from '@microsoft/sp-core-library';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PeoplePickerItem } from 'office-ui-fabric-react';
import { SPOperations } from '../Services/SPOps'

/*export interface IPeople {
    DepartmentId: number;
    Jobtype: string;
    ManagerId: number;
    Title: string;
    Lastname: string;
    Role: string;
    Id: number;
}
*/
export class People{
    DepartmentId: number;
    Jobtype: string;
    ManagerId: number;
    Title: string;
    Lastname: string;
    Role: string;
    Id: number;
    
    constructor(id, departmentId, jobtype, managerId, lastname, role, title) {
        this.DepartmentId = departmentId;
        this.Jobtype = jobtype;
        this.ManagerId = managerId;
        this.Title = title;
        this.Lastname = lastname;
        this.Role = role;
        this.Id = id;
    }
}


//<input type="" name="ManagerId" value={this.state.ManagerId} onChange={event => this.handleChange(event)} />
export interface ICrudState {
    data: any,
    
    Title: string,
    Lastname: string,
    Jobtype: string,
    Role: string,
    ManagerId: string,
    DepartmentId: string,
    presentId: string,
    roleChoices: any,
    JobtypeChoice: any,
    newValue: any,


} 

export default class PnpCrud extends React.Component<IPnpCrudProps, ICrudState> {
    employeeService: SPOperations = new SPOperations();;
  
    constructor(props) {
        super(props);
        this.state = {
            data:[],
            Title: '',
            Lastname: '',
            Jobtype: '',
            Role: '',
            ManagerId: '',
            DepartmentId: '',
            presentId: '',
            roleChoices: [],
            JobtypeChoice: [],
            newValue: [],
       
        }
        this.createItem = this.createItem.bind(this);
        this.DeleteItem = this.DeleteItem.bind(this);
        this.EditItem = this.EditItem.bind(this);
        this.handleChange = this.handleChange.bind(this);
        this.EditNew = this.EditNew.bind(this);
      
        this.handlePeople = this.handlePeople.bind(this);
    }
    async createItem() {
        var newItem = {
            Title: this.state.Title,
            Lastname: this.state.Lastname,
            Jobtype: this.state.Jobtype,
            DepartmentId: this.state.DepartmentId,
            Role: this.state.Role,
            ManagerId: this.state.newValue.Id
        }

        console.log(this.state);
        await sp.web.lists.getByTitle('Employee OnBoard').items.add(newItem);
        this.getItems();
    }

    async handlePeople(items:any) {
        console.log(items[0]);
        const user = await sp.web.siteUsers.getByLoginName(items[0].loginName).get();
        console.log(user);
        const user2 = await sp.web.siteUsers.getByEmail(user.Email).get();
        console.log(user2);
        this.setState({
            newValue:user2
        })
    }

    async DeleteItem(Id) {
        await sp.web.lists.getByTitle('Employee OnBoard').items.getById(Id).delete();
        this.getItems();
    }
    
    EditItem(Id) {
        var x;
        this.state.data.forEach(element => {
            if (element.Id == Id) {
                x = element;
            }
        });
   
        this.setState({
            Title: x.Title,
            DepartmentId: x.DepartmentId,
            Lastname: x.Lastname,
            ManagerId: x.ManagerId,
            Role: x.Role,
            Jobtype: x.Jobtype
        });
        this.setState({
            presentId: Id
        });
        
    }
    
    async EditNew() {
        var id: number = Number(this.state.presentId);
        var newItem = {
            Title: this.state.Title,
            Lastname: this.state.Lastname,
            Jobtype: this.state.Jobtype,
            DepartmentId: this.state.DepartmentId,
            Role: this.state.Role,
            ManagerId: this.state.ManagerId
        }
        await sp.web.lists.getByTitle('Employee OnBoard').items.getById(id).update(newItem);
        this.getItems();
        

    }
    handleChange(e) {
        let change = {};
        change[e.target.name] = e.target.value;
        this.setState(change);
        
    }

    async componentDidMount() {
        this.getItems();
        const items:any = await sp.web.lists.getByTitle('Employee OnBoard').fields.getByInternalNameOrTitle('Role').select('Choices,ID').get();
       
        const jt: any = await sp.web.lists.getByTitle('Employee OnBoard').fields.getByInternalNameOrTitle('Jobtype').select('Choices,ID').get();
        this.setState({
            JobtypeChoice: jt.Choices,
            roleChoices: items.Choices
        })
    }
 
   
    async getItems() {
        


        const items = await sp.web.lists.getByTitle('Employee OnBoard').items.getAll();
        console.log(items)
        let newLog = [];
        items.forEach(async (item) => {

            var user = await sp.web.siteUsers.getById(item.ManagerId).get();
            item['user'] = user;
            newLog.push(item);
            this.setState({
                data: newLog
            });
            
        });
        console.log(newLog)
        
    }
   

  public render(): React.ReactElement<IPnpCrudProps> {
      return (
          <div className={styles.pnpCrud}>
              <div className={styles.container}>
                  <div className={styles.row}>
                      <div className={styles.column}>
                        <span className={styles.title} >Welcome to Employee OnBoard (PNPjs)</span>
                            <table>
                                <tr>
                                    <th>First Name</th>
                                    <th>Last Name</th>
                                    <th>Job Type</th>
                                    <th>Department</th>
                                    <th>Role</th>
                                    <th>Manager</th>
                                </tr>

                                {this.state.data.map((d: any) => (
                                    <tr>
                                        <td onClick={i => this.EditItem(d.Id)}>{d.Title}</td>
                                        <td>{d.Lastname}</td>
                                        <td>{d.Jobtype}</td>
                                        <td>{d.DepartmentId}</td>
                                        <td>{d.Role}</td>
                                        <td>{d.user.Title}</td>
                                        <td><button onClick={i => this.DeleteItem(d.Id)}>DELETE</button></td>
                                    </tr>
                                ))}
                                <tr>
                                    <td><input type="" name="Title" value={this.state.Title} onChange={event => this.handleChange(event)} /></td>
                                    <td><input type="" name="Lastname" value={this.state.Lastname} onChange={event => this.handleChange(event)} /></td>
                                    
                                  <td>
                                      <select onChange={event => this.handleChange(event)} name="Jobtype" value={this.state.Jobtype}>
                                              {this.state.JobtypeChoice.map((jt) => <option value={jt}>{jt}</option>)}
                                          </select>
                                      </td>

                                    <td><input type="" name="DepartmentId" value={this.state.DepartmentId} onChange={event => this.handleChange(event)} /></td>

                                  <td>
                                      <select onChange={event => this.handleChange(event)} name="Role" value={this.state.Role}>
                                          {this.state.roleChoices.map((rc) => <option value={rc}>{rc}</option>)}
                                          </select>
                                      </td>
                                  
                                   <td><PeoplePicker titleText={"Employee Name"} placeholder="Enter" onChange={this.handlePeople} context={this.props.context}></PeoplePicker></td>
                                </tr>

                          </table>
                          
                            <button onClick={this.createItem}>Create</button>
                            <button onClick={this.EditNew}>Edit</button>
                      </div>
                  </div>
               </div>
          </div>
    );
  }
}
