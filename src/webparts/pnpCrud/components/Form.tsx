import * as React from 'react';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import styles from './PnpCrud.module.scss';
import { IPnpCrudProps } from './IPnpCrudProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Log } from '@microsoft/sp-core-library';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PeoplePickerItem, PrimaryButton } from 'office-ui-fabric-react';
import { SPOperations } from '../Services/SPOps';
import {IUnderstandStateComponentProps} from './PnpCrud';
import { TextField, MaskedTextField } from '@fluentui/react/lib/TextField';
import {IFormProps} from './PnpCrud';
import {useState} from 'react';
import { useId, useBoolean } from '@fluentui/react-hooks';
import {
  getTheme,
  mergeStyleSets,
  FontWeights,
  ContextualMenu,
  Toggle,
  Modal,
  IDragOptions,
  IIconProps,
  Stack,
  IStackProps,
} from '@fluentui/react';
import { DefaultButton, IconButton, IButtonStyles } from '@fluentui/react/lib/Button';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';

 
const dragOptions: IDragOptions = {
      moveMenuItemText: 'Move',
      closeMenuItemText: 'Close',
      menu: ContextualMenu,
    };

    const cancelIcon: IIconProps = { iconName: 'Cancel' };
    const dropdownStyles: Partial<IDropdownStyles> = {
      title:{color:'red'},
      dropdown: { width: 300 },
    };

    
    export const Form: React.FunctionComponent<IFormProps> = (props) => {
      var [isModalOpen, { setTrue: showModal, setFalse: hideModal }] = useBoolean(false);
      const [isDraggable, { toggle: toggleIsDraggable }] = useBoolean(false);
   
      
      var jt=props.Jobtype;
      var role=props.Role;
      var  dep="";
      var user="";
      console.log('hello');
      const bntRef=React.useRef(null);
      console.log(props.defaultUser);
      if(props.b){
        isModalOpen=true;
        //showModal();  
     
        
      }
      function hide(){
        //isModalOpen=false;
        //console.log('hello');
        hideModal();
        props.b=false;
        this.forceUpdate();
        
      }
    
      // Use useId() to ensure that the IDs are unique on the page.
      // (It's also okay to use plain strings and manually ensure uniqueness.)
      const titleId = useId('title');
      function handleSubmit(event:any,item){
            console.log(item);
            event.preventDefault();
          }

      function handleSub(event){
            event.preventDefault();
            var response={
              Title:event.target.Title.value,
              Lastname:event.target.LastName.value,
              Department:dep,
              Role:role,
              Jobtype:jt,
              loginName:user
            }
            console.log(event.target.Title.value);
            console.log(jt);
            hideModal();
            props.CreateFunction(response);
            
        }
      function handleJt(event,item){
            jt=item.key;
            console.log(jt);
            console.log(event);
            
      }
      function handleDep(event,item){
          dep=item.text;

      }
      function handleRole(event,item){
          role=item.key;
      }
      function handlePeople(items){
        user=items[0].loginName;
      }
      

      return (
        <div>
          
          <DefaultButton secondaryText="Opens the Sample Modal" onClick={showModal} text="Create New Item"  ref={bntRef}/>
          <Modal
            titleAriaId={titleId}
            isOpen={isModalOpen}
            onDismiss={hideModal}
            isModeless={true}
            containerClassName={contentStyles.container}
            dragOptions={isDraggable ? dragOptions : undefined}
          >
            <div className={contentStyles.header}>
              <span id={titleId}>Employee Onboard Form</span>
              <IconButton
                styles={iconButtonStyles}
                iconProps={cancelIcon}
                ariaLabel="Close popup modal"
                onClick={hideModal}
              />
            </div>
    
            <div className={contentStyles.body}>
             <form onSubmit={handleSub}>
                   <TextField label="First Name" name="Title" defaultValue={props.Title}/>
                   <TextField label="Last Name" name="LastName" defaultValue={props.Lastname} />
                   <Dropdown placeholder="Select an option" label="JobType" options={props.jobChoices}  styles={dropdownStyles} defaultSelectedKey={props.Jobtype} onChange={handleJt}/>
                   <Dropdown placeholder="Select an option" label="Department" options={props.depChoices} styles={dropdownStyles}  defaultSelectedKey={props.depId} onChange={handleDep}/>
                   <Dropdown placeholder="Select an option" label="Role" options={props.rChoices} styles={dropdownStyles} defaultSelectedKey={props.Role} onChange={handleRole}/>
                   <PeoplePicker titleText={"Employee Name"} placeholder="Enter" context={props.context} onChange={handlePeople} defaultSelectedUsers={props.defaultUser}></PeoplePicker>
                   <PrimaryButton  text="Submit" type="submit" />

             </form>

            </div>
          </Modal>
        </div>
      );
    };
    
    const theme = getTheme();
    const contentStyles = mergeStyleSets({
      container: {
        display: 'flex',
        flexFlow: 'column nowrap',
        alignItems: 'stretch',
      },
      header: [
        // eslint-disable-next-line deprecation/deprecation
        theme.fonts.xLargePlus,
        {
          flex: '1 1 auto',
          borderTop: `4px solid ${theme.palette.themePrimary}`,
          color: theme.palette.neutralPrimary,
          display: 'flex',
          alignItems: 'center',
          fontWeight: FontWeights.semibold,
          padding: '12px 12px 14px 24px',
        },
      ],
      body: {
        flex: '4 4 auto',
        padding: '0 24px 24px 24px',
        overflowY: 'hidden',
        selectors: {
          p: { margin: '14px 0' },
          'p:first-child': { marginTop: 0 },
          'p:last-child': { marginBottom: 0 },
        },
      },
    });
    const toggleStyles = { root: { marginBottom: '20px' } };
    const iconButtonStyles = {
      root: {
        color: theme.palette.neutralPrimary,
        marginLeft: 'auto',
        marginTop: '4px',
        marginRight: '2px',
      },
      rootHovered: {
        color: theme.palette.neutralDark,
      },
    };
    