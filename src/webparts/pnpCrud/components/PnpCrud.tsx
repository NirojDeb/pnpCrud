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
    data: People[],
    Title: string,
    Lastname: string,
    Jobtype: string,
    Role: string,
    ManagerId: string,
    DepartmentId: string,
    presentId: string,
    roleChoices: string[],
    JobtypeChoice: string[],
    newValue: any,
    DepartmentChoices: {},
    Dep: string[],
    revMap: {}

} 

export default class PnpCrud extends React.Component<IPnpCrudProps, ICrudState> {
    employeeService: SPOperations = new SPOperations();
  
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
            DepartmentChoices: {},
            Dep: [],
            revMap: {}
        }

        this.createItem = this.createItem.bind(this);
        this.DeleteItem = this.DeleteItem.bind(this);
        this.EditItem = this.EditItem.bind(this);
        this.handleChange = this.handleChange.bind(this);
        this.EditNew = this.EditNew.bind(this);
        this.handlePeople = this.handlePeople.bind(this);
    }

    private async createItem(): Promise<any> {
        
        console.log(this.state.DepartmentChoices);



        let e = new People(1, this.state.DepartmentId, this.state.Jobtype, this.state.newValue.Id, this.state.Lastname, this.state.Role, this.state.Title);
        await this.employeeService.CreateItem(e);
        
        this.getItems();
    }

    private  async handlePeople(items:any):Promise<any> {
      
        const user = await this.employeeService.GetUserByLoginName(items[0].loginName);
        this.setState({
            newValue: user
        });
    }

    private  async DeleteItem(Id):Promise<any> {
        await this.employeeService.DeleteItem(Id);
        this.getItems();
    }
    
    private  EditItem(Id):void  {
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
    
    private  async EditNew() {
        var id: number = Number(this.state.presentId);
        let e: People = new People(1, this.state.DepartmentId, this.state.Jobtype, this.state.newValue.Id, this.state.Lastname, this.state.Role, this.state.Title);
        await this.employeeService.EditItem(id, e);
        this.getItems();
  
    }
    private  handleChange(e):void {
        let change = {};
        change[e.target.name] = e.target.value;
        this.setState(change);
        
    }

    async componentDidMount():Promise<any> {
        this.getItems();
        const roles: any = await this.employeeService.GetRoles();
        const jt: any = await this.employeeService.GetJobTypes();
        const res: [] = await this.employeeService.GetDepartements();
        let x = {};
        let y = [];
        let z = {}
        res.forEach((r) => {
            x[r[0]] = r[1];
            y.push(r[0]);
            z[r[1]] = r[0];

        })
        
        this.setState({
            JobtypeChoice: jt.Choices,
            roleChoices: roles.Choices,
            DepartmentChoices: x,
            Dep: y,
            revMap:z
        });
        
    }
 
   
    private async getItems(): Promise<any> {
        const items = await this.employeeService.GetAllItems();
        let newLog = [];
        await items.forEach(async (item) => {

            var user = await this.employeeService.GetUser(item.ManagerId);
            item['user'] = user;
            newLog.push(item);
            this.setState({
                data: newLog
            });
            
        });
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
                                        <td>{this.state.revMap[d.DepartmentId]}</td>
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

                                  <td>
                                      <select onChange={event => this.handleChange(event)} name="DepartmentId" value={this.state.DepartmentId}>
                                          {this.state.Dep.map((dep) => <option value={this.state.DepartmentChoices[dep]}>{dep}</option>)}
                                      </select>
                                    </td>

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
