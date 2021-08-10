import * as React from 'react';
import styles from './SpfxWebPartDemoNj.module.scss';
import { ISpfxWebPartDemoNjProps } from './ISpfxWebPartDemoNjProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { AuditService } from '../Services/AuditSetvice';

interface IAudit{
  Title?:String,
  Id?:Number,
  Name:String,
  Sponsor:String,
  Location:String,
  Deadline:String,
  TotalTasks:Number,
  AssignedTo:String,
  AttachmentsArray?:[]

}

export interface IAuditState{
  Audits: IAudit[],
  presentAttachments:any
}

export default class SpfxWebPartDemoNj extends React.Component<ISpfxWebPartDemoNjProps, IAuditState> {
  auditService:AuditService=new AuditService();
  FileInfos:any=[];

  constructor(props)
  {
    
    super(props);
    this.state={
      Audits:[],
      presentAttachments:[]
    }
    var FileInfos:any=[];
    this.downloadAttachment=this.downloadAttachment.bind(this);
    this.submitForm=this.submitForm.bind(this);
    this.setAttachments=this.setAttachments.bind(this);
  }

  public downloadAttachment(attachment)
  {
    let url=`https://niroj.sharepoint.com${attachment.ServerRelativeUrl}`;
    window.open(url);
  }

  async submitForm(event)
  {
 
    let newAudit={
      Title:'Hello',

      Name:event.target.Name.value,
      Sponsor:event.target.Sponsor.value,
      Location:event.target.Location.value,
      TotalTasks:event.target.TotalTasks.value,
      Deadline:'2015-08-11T12:40:16Z',
      AssignedToId:11
    }

    let createdAudit=await this.auditService.createAudit(newAudit);
    console.log(createdAudit);
    await this.auditService.addAttachementsToAudit(createdAudit.data.Id,this.FileInfos);
    
    var audits= await this.auditService.getListItems();
    
    

    

    this.setState({
      Audits:audits
    });
    console.log(this.FileInfos);
    
    
  }

   filereader(file)
  {
    return new Promise((resolve,reject)=>{
      const reader = new FileReader();
      
    })
  }

  async setAttachments(event:any)
  {
    const _that=this;

    var fileCount = event.target.files.length;
   
    for (var i = 0; i < fileCount; i++) {
       var fileName = event.target.files[i].name;
       console.log(fileName);
       var file = event.target.files[i];
       var reader = new FileReader();
       reader.onload = (function(file) {
          return async function(e) {
             
             //Push the converted file into array
                _that.FileInfos.push({
                   "name": file.name,
                   "content": e.target.result
                   });
                
                }
          })(file);
       reader.readAsArrayBuffer(file);
     }
 
    //  this.setState({
    //    presentAttachments:fileInfos
    //  })
  }

  async componentDidMount(){
    let res=await this.auditService.getListItems();
    let resAudits:IAudit[]=[]
    for(var audit of res)
    {
      let attachments=await this.auditService.getAttachmentsbyId(audit.Id);

      let resAudit:IAudit={
        Id:audit.Id,
        Name:audit.Name,
        Location:audit.Location,
        AssignedTo:audit.AssignedToId,
        Deadline:audit.Deadline,
        TotalTasks:audit.TotalTasks,
        Sponsor:audit.Sponsor,
        AttachmentsArray:attachments.AttachmentFiles
      }


      resAudits.push(resAudit);
    }
    console.log(res);
    
    this.setState({
      Audits:resAudits
    })
  }


  public render(): React.ReactElement<ISpfxWebPartDemoNjProps> {
    return (
      <div className={ styles.spfxWebPartDemoNj }>
        <div className={ styles.container }>
          <div>
            <table>
              <tr>
                <th>ID</th>
                <th>Name</th>
                <th>Sponsor</th>
                <th>Location</th>
                <th>Deadline</th>
                <th>Total Tasks</th>
                <th>Assigned To</th>
                <th>Attachments</th>
              </tr>
              {this.state.Audits.map((audit:IAudit)=>{
                return (
                <tr>
                  <td>{audit.Id}</td>
                  <td>{audit.Name}</td>
                  <td>{audit.Sponsor}</td>
                  <td>{audit.Location}</td>
                  <td>{audit.Deadline}</td>
                  <td>{audit.TotalTasks}</td>
                  <td>{audit.AssignedTo}</td>
                  <td>
                  {audit.AttachmentsArray!=undefined?<div>{audit.AttachmentsArray.map((attachment:any)=>{
                    return(
                      <li onClick={()=>{this.downloadAttachment(attachment)}}>{attachment.FileName}</li>
                    )
                  })}</div>:''}
                  
                  </td>
                  
                </tr>)
              })}
            </table>
            <form onSubmit={this.submitForm}>
              <label>Name &nbsp;&nbsp;</label>
              <input type="text" name="Name" />
              <br></br>
              <label>Sponsor &nbsp;&nbsp;</label>
              <input type="text" name="Sponsor" />
              <br></br>
              <label>Location &nbsp;&nbsp;</label>
              <input type="text" name="Location" />
              <br></br>
              <label>Deadline &nbsp;&nbsp;</label>
              <input type="text" name="Deadline" />
              <br></br>
              <label>Total Tasks &nbsp;&nbsp;</label>
              <input type="number" name="TotalTasks" />
              <br></br>
              <label>Attachment &nbsp;&nbsp;</label>
              <input type="file" multiple name="Attachments" onChange={this.setAttachments} />
              <br></br>
              <button type="submit">Submit</button>
            </form>
          </div>
        </div>
      </div>
    );
  }
}
