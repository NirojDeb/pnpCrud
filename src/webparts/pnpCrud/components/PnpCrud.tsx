import * as React from 'react';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import styles from './PnpCrud.module.scss';
import { IPnpCrudProps } from './IPnpCrudProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Log } from '@microsoft/sp-core-library';

export interface ICrudState {
    data: any,
    filler: any,
    Title: string,
    Lastname: string,
    Jobtype: string,
    Role: string,
    ManagerId: string,
    DepartmentId: string,
    presentId: string,
    roleChoices: any,
    JobtypeChoice:any

}

export default class PnpCrud extends React.Component<IPnpCrudProps, ICrudState> {
    constructor(props) {
        super(props);
        this.state = {
            data: [],
            filler: [],
            Title: '',
            Lastname: '',
            Jobtype: '',
            Role: '',
            ManagerId: '',
            DepartmentId: '',
            presentId: '',
            roleChoices: [],
            JobtypeChoice:[]
        }
        this.createItem = this.createItem.bind(this);
        this.DeleteItem = this.DeleteItem.bind(this);
        this.EditItem = this.EditItem.bind(this);
        this.handleChange = this.handleChange.bind(this);
        this.EditNew = this.EditNew.bind(this);
        this.handleRoleDropdown = this.handleRoleDropdown.bind(this);
    }
    async createItem() {
        var newItem = {
            Title: this.state.Title,
            Lastname: this.state.Lastname,
            Jobtype: this.state.Jobtype,
            DepartmentId: this.state.DepartmentId,
            Role: this.state.Role,
            ManagerId: this.state.ManagerId
        }

        console.log(this.state);
        await sp.web.lists.getByTitle('Employee OnBoard').items.add(newItem);
        this.getItems();
        

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
        console.log(x);
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
        console.log(this.state);
        console.log(x);
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
        console.log(e.target.name);
        console.log(e.target.value);
        console.log(this.state.Role);
        console.log(this.state.Jobtype);
    }
        
    async componentDidMount() {
        this.getItems();
        const items:any = await sp.web.lists.getByTitle('Employee OnBoard').fields.getByInternalNameOrTitle('Role').select('Choices,ID').get();
        
        this.setState({
            roleChoices:items.Choices
        })

        const jt: any = await sp.web.lists.getByTitle('Employee OnBoard').fields.getByInternalNameOrTitle('Jobtype').select('Choices,ID').get();
        this.setState({
            JobtypeChoice:jt.Choices
        })
    }


    async getItems() {
        console.log('dwdwdwd');
        
        const items: any[] = await sp.web.lists.getByTitle('Employee OnBoard').items.getAll();
        
        this.setState({
            data: items
        });
        console.log(this.state.data);
    }
    handleRoleDropdown(e) {
        console.log(e.target.value);
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
                                        <td>{d.ManagerId}</td>
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
                                  
                                    <td><input type="" name="ManagerId" value={this.state.ManagerId} onChange={event => this.handleChange(event)} /></td>
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
