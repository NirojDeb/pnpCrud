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
import { SPOperations } from '../Services/EmployeeService'
import  {Form}  from './Form';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';

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
export interface IUnderstandStateComponentProps {
    title: string;
}

export interface IFormProps{
   
    context:WebPartContext;
    Title: string,
    Lastname: string,
    Jobtype: string,
    Role: string,
    jobChoices:IDropdownOption[],
    rChoices:IDropdownOption[],
    depChoices:IDropdownOption[],
    depId:string,
    CreateFunction(any):void,
    b:boolean,
    defaultUser:string[]
    
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
    revMap: {},
    jtChoices:IDropdownOption[],
    depChoices:IDropdownOption[],
    rChoices:IDropdownOption[],
    mode:string,
    toggle:boolean,
    presentUser:any

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
            revMap: {},
            jtChoices:[],
            rChoices:[],
            depChoices:[],
            mode:'',
            toggle:false,
            presentUser:[]

        }

        this.createItem = this.createItem.bind(this);
        this.DeleteItem = this.DeleteItem.bind(this);
        this.EditItem = this.EditItem.bind(this);
       
        this.EditNew = this.EditNew.bind(this);
        
        this.CreateFunction=this.CreateFunction.bind(this);
        const form = React.createRef();
    }

    private async createItem(): Promise<any> {
        
        console.log(this.state.DepartmentChoices);



        let e = new People(1, this.state.DepartmentId, this.state.Jobtype, this.state.newValue.Id, this.state.Lastname, this.state.Role, this.state.Title);
        await this.employeeService.CreateItem(e);
        
        this.getItems();
        
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
        console.log(x.user);
        var y=[x.user.LoginName];
        console.log(x);
        
        this.setState({
            Title: x.Title,
            DepartmentId: x.DepartmentId,
            Lastname: x.Lastname,
            ManagerId: x.ManagerId,
            Role: x.Role,
            Jobtype: x.Jobtype,
            presentUser:y,
            presentId: Id,
            mode:'edit',
            toggle:true
        });
        
    
        
    }
    
    private  async EditNew() {
        var id: number = Number(this.state.presentId);
        let e: People = new People(1, this.state.DepartmentId, this.state.Jobtype, this.state.newValue.Id, this.state.Lastname, this.state.Role, this.state.Title);
        await this.employeeService.EditItem(id, e);
        this.getItems();
        this.setState({
            mode:'create'
        });

  
    }    
 
   
    private async getItems(): Promise<any> {
        const items = await this.employeeService.GetAllItems();
        const newLog = await this.NewFunc(items);
      
        
        this.setState({
            data: newLog
        });
        
        console.log(newLog);
    }

    public async NewFunc(items: any): Promise<any> {

        let newLog = [];
        for(let item of items) {

            var user = await this.employeeService.GetUser(item.ManagerId);
            item['user'] = user;
            newLog.push(item);
        }
        return newLog;

    }

    public async componentDidMount(): Promise<any> {
        this.getItems();
        const roles: any = await this.employeeService.GetRoles();
        const jt: any = await this.employeeService.GetJobTypes();
        const res: [] = await this.employeeService.GetDepartements();
        let x = {};
        let y = [];
        let z = {};

        let q:any=[];
        let w:any=[];
        let e:any=[];

        for(var jb of jt.Choices)
        {
            var temp:IDropdownOption={
                key:jb,
                text:jb
            };
            console.log(temp);
            
            q.push(temp);
        }

        for (var r of roles.Choices)
        {
            var temp:IDropdownOption={
                key:r,
                text:r
            }
            w.push(temp);
        }

        for (var p of res)
        {
            var temp:IDropdownOption={
                key:p[1],
                text:p[0],
            }
            e.push(temp);
        }

     
        

        res.forEach((r) => {
            x[r[0]] = r[1];
            y.push(r[0]);
            z[r[1]] = r[0];

        });
        console.log(q);
        

        this.setState({
            JobtypeChoice: jt.Choices,
            roleChoices: roles.Choices,
            DepartmentChoices: x,
            Dep: y,
            revMap: z,
            jtChoices:q,
            rChoices:w,
            depChoices:e,
            mode:'create'
        });
        console.log(this.state.DepartmentChoices)

    }

    public async  CreateFunction(value){
        console.log(value);
        console.log('heeeee');
        console.log(value.Department);
        
        var d=value.Department.toString();
        console.log(d);
        
        var depId=this.state.DepartmentChoices[d];
        const user = await this.employeeService.GetUserByLoginName(value.loginName);
        this.setState({
            Title:value.Title,
            DepartmentId:depId,
            Lastname:value.Lastname,
            Role:value.Role,
            Jobtype:value.Jobtype,
            newValue:user,
            toggle:false

        });
        if(this.state.mode=='create')
        {
            this.createItem();
        }
        else{
            this.EditNew();
        }
        
        
        
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
                                

                          </table>
                          
                            
                      </div>
                  </div>
              </div>
           <Form CreateFunction={this.CreateFunction} Title={this.state.Title} depId={this.state.DepartmentId} Lastname={this.state.Lastname} Jobtype={this.state.Jobtype}  Role={this.state.Role} context={this.props.context} jobChoices={this.state.jtChoices} rChoices={this.state.rChoices} depChoices={this.state.depChoices} b={this.state.toggle} defaultUser={this.state.presentUser}/>
          </div>
    );
  }
}
