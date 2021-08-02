import {CreateClientsidePage, PrincipalSource, PrincipalType, sp, Web} from '@pnp/sp/presets/all';
import {WebPartContext} from '@microsoft/sp-webpart-base';

export class SPoperations 
{
    
     constructor(context:WebPartContext)
     {
        sp.setup({  
            spfxContext: context
        });  
     }
    
     public getListTitle() : Promise<any[]>
     {
          let listTitle : any[] = [];
          console.log(sp);
          return new Promise<any[]>(async(resolve,reject)=>{

                sp.web.lists.select("Title").get().then(
                    (results:any)=>{
                      //console.log(results);
                      results.map((result:any)=>{
                          listTitle.push({key:result.Title,value:result.Title})
                      })

                      resolve(listTitle);
                   }
                   ,(error:any)=>{
                        console.log(error);
                        reject("error occured");
                   }
                );
          })
     }

     public getattachementDetails() : Promise<void>
     {
           return new Promise<void>(async(resolve,reject)=>{

                   try{

                          let item = await sp.web.lists.getByTitle("samplelist").items.getById(1);
                          //console.log(item);

                          let files = await item.attachmentFiles();
                          //console.table(files);


                          resolve() 
                   }
                   catch{

                        reject("error")
                   }
           })
     }
     public getPageDetails() : Promise<void>
     {
           return new Promise<void>(async(resolve,reject)=>{

                    try{

                        /*
                        // const page = await sp.web.addClientsidePage("MyPageTest1","My Page For Testing Purpose");
                        // await page.save();

                        // const page2 = await CreateClientsidePage(sp.web,"MyPageTest2","My Page for Testing Purpose 2")
                        //page2.save();
                         */

                         /*
                         
                         const spweb = Web("https://pakur2.sharepoint.com/sites/moderncommunicationsite");
                         //const r = await spweb();
                         
                         const page = await spweb.addClientsidePage("Mypage1test","My page title");
                         await page.save();*/

                         /* 
                         
                         const spweb = Web("https://pakur2.sharepoint.com/sites/moderncommunicationsite");

                         const list = await spweb.lists.getByTitle("Large List");

                         const views = await list.views();

                         views.forEach((item)=>{
                             console.log("View Title - " + item.Title);
                         })
                         */

                         /*  
                         const spweb = Web("https://pakur2.sharepoint.com/sites/moderncommunicationsite");
                         const userCustomActions = await spweb.userCustomActions();

                         userCustomActions.forEach((item)=>{
                             console.log("Custom Actions - " + item.Title)
                         })
                         */

                         /* 

                         const termstore = await sp.termStore();
                         console.log("Term Store ID - " + termstore.id);
                         console.log("Term Store name - " + termstore.name);
                          
                         const termstoregroups = await sp.termStore.groups();

                         for(const item of termstoregroups){
                             console.log("Term Store Group " + item.name);
                             console.log("Term Store Group id - " + item.id);

                             const termsets = await sp.termStore.groups.getById(item.id).sets();
                             
                             for(const termset of termsets)
                             {
                                console.log("Term Set id - " + termset.id);
                                const termsetinfo = await sp.termStore.sets.getById(termset.id)();

                                termsetinfo.localizedNames.map((item)=>{ 
                                    console.log("Term Set localized Name -" + item.name);
                                });

                                const terms = await sp.termStore.sets.getById(termset.id).terms();
                                for(const term of terms)
                                {
                                    console.log("Term - " + term.id)

                                    term.labels.map((item)=>{ 
                                        console.log("Term - " + item.name)
                                    });
                                    
                                }

                             }


                             
                         }
                         */
                        /*  
                        let currentuseremail = await sp.utility.getCurrentUserEmailAddresses();
                        console.log("Email Address - " + currentuseremail)

                        let principalinfo = await sp.utility.resolvePrincipal("raji@pakur2.onmicrosoft.com",PrincipalType.User,PrincipalSource.All,true,false)
                        console.log("Principal Email - " + principalinfo.Email)
                        console.log("Principal LoginName - " + principalinfo.LoginName)
                        console.log("Principal PrincipalId - " + principalinfo.PrincipalId)
                        console.log("Principal PrincipalType - " + principalinfo.PrincipalType)
                        console.log("Principal SIPAddress - " + principalinfo.SIPAddress)
                         
                    
                        //sp.utility.searchPrincipals()
                        //sp.utility.createEmailBodyForInvitation()

                        const emailProps  = {
                            To: ["me.somnath.singh@gmail.com"],
                            CC: ["somnath@pakur2.onmicrosoft.com"],
                            Subject: "This email is about...",
                            Body: "Here is the body. <b>It supports html</b>",
                            AdditionalHeaders: {
                                "content-type": "text/html"
                            }
                        };
                        
                        await sp.utility.sendEmail(emailProps);
                        */
                        /* 
                         let followedsite = await sp.social.getFollowedSitesUri();
                         console.log("Followed Site - " + followedsite)

                         let followeddocuments = await sp.social.getFollowedDocumentsUri();
                         console.log("Followed Site - " + followeddocuments)
                         */
                        /* 
                        const username = "bubai@pakur2.onmicrosoft.com";
                        let user = false;
                        try{
                            let result = await sp.web.ensureUser(username);
                            user = true;
                        }
                        catch{
                            user = false
                        }
                        
                        console.log("Ensure User - "+ user);
                        */

                        /*
                        const web = Web("https://pakur2.sharepoint.com/sites/ClassicSite");

                        const item = await web.lists.getByTitle("Employee").items.getById(1);

                        await item.resetRoleInheritance();
                        await item.breakRoleInheritance(false,false);


                         
                        const user = await web.siteUsers.getByEmail("raji@pakur2.onmicrosoft.com").get();
                        console.log("User ID - " + user.Id);

                        await item.roleAssignments.add(user.Id,1073741826)

                        const currentuser = await web.currentUser();
                        await item.roleAssignments.remove(currentuser.Id,1073741829)
                       
                        //await item.resetRoleInheritance();
                        */
                        console.log("Hello World");
                        const web = Web("https://pakur2.sharepoint.com/sites/ClassicSite");
                        const list = await web.lists.getByTitle("Employee");
                        const items = await list.items.filter("Title eq 'Mr' or startswith(City,'City')").get();

                        for(const item of items)
                        {
                            //console.log(item);
                            console.log(item.Title);
                            console.log(item.Name);
                            console.log(item.City);
                            console.log(item.Department);
                            console.log("-------------------------------")
                        }
                        
                        let item = {
                            Title:"Mr",
                            Name:"S Singh",
                            City:"City 1",
                            Department:"Dept D",
                            Employee_x0020_Code:"E10615"
                            
                        }
                        const additem = list.items.add(item);
                        console.log(additem);
                         resolve() 
                        }
                        catch(ex){
                             console.log(ex)
                             reject("error")
                        }
           })
     }

}