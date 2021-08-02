import * as React from 'react';
import styles from './SpfxPnpjx.module.scss';
import { ISpfxPnpjxProps } from './ISpfxPnpjxProps';
import { ISpfxPnpjxState } from './ISpfxPnpjxState'

import {SPoperations} from '../services/SPoperations';

import { escape } from '@microsoft/sp-lodash-subset';

import {Dropdown, PrimaryButton} from 'office-ui-fabric-react';
import {sp} from '@pnp/sp/presets/all';

export default class SpfxPnpjx extends React.Component<ISpfxPnpjxProps,ISpfxPnpjxState, {}> {

  private _spService:SPoperations;

  constructor(props:ISpfxPnpjxProps)
  {
     super(props);
     this._spService = new SPoperations(this.props.spcontext);
     this.state = {
        listTitle : []
     };
     this._spService.getattachementDetails();
     this._spService.getPageDetails();
    
  }
  
  public componentDidMount()
  {
     this._spService.getListTitle().then((result)=>{
        
         this.setState({listTitle:result});

     })
  } 

 
  
  
  public render(): React.ReactElement<ISpfxPnpjxProps> {
    
    let options = [];
    
    if(this.state.listTitle.length > 0)
    {
      options.push(<option value="-1">-- Select List --</option>)
        this.state.listTitle.map((result)=>{
                options.push(<option value={result.key}>{result.value}</option>);
       })

     
    }
    

    return (
      <div className={ styles.spfxPnpjx }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
            </div>
            <div className={ styles.column }>
                 <select>
                   {
                     options
                   }
               
                 </select>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
