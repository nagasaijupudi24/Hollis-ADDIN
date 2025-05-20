/* eslint-disable prettier/prettier */
// import "./excel";
// import "./outlook";
// import "./powerpoint";
// import "./word";


/* eslint-disable no-case-declarations */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
/* eslint-disable @typescript-eslint/no-unused-vars */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Office */

import { AccountManager } from "./authConfig";
// import { makeGraphRequest } from "./msgraph-helper";
// import { writeFileNamesToOfficeDocument } from "./document";

const accountManager = new AccountManager();
let selectedProjectIdTable;
let selectedProjectTaskIdTable;
let choiceOptions: any[];

let selectedProjectIdNew: any;
let selectedProjectName: any;
let selectedProjectTaskIdNew: any;
let selectedProjectTaskName: any;

let pca = undefined;
let isPCAInitialized = false;
let token: any;

let domain:any = "https://hollis-projectops-uat-01.crm4.dynamics.com";
// https://hollis-projectops-dev-01.api.crm4.dynamics.com

let options: any[] = [];
let dropdownListProject:any = document.getElementById("dropdownListForProject");
let dropdownListTask:any = document.getElementById("dropdownListForTask");
let searchInputProject:any = document.getElementById("searchInputForProject");
let searchInputTask:any = document.getElementById("searchInputForTask");
let projectError:any = document.getElementById("projectError");
let taskError:any = document.getElementById("taskError");
let duration:any = document.getElementById("duration")
let durationError:any = document.getElementById("durationError")
let insertButton:any = document.getElementById("insertTimeEntry")
let insertError:any = document.getElementById("insertError");
let date:any = document.getElementById("date")
let dateError:any = document.getElementById("dateError")
let description:any = document.getElementById("description")
let descriptionTextarea:any = document.getElementById("descriptionTextarea")
let DescriptionError:any = document.getElementById("DescriptionError")

let projectnameArray:any = [];
let projectTaskArr:any = [];

let projectType:any = "Client"


let client:any = document.getElementById("client")
let internal:any = document.getElementById("internal")
client.addEventListener('click',()=>{
  // console.log("client")
  
  client.classList.add("active", "toggle-btn")
  internal.classList.remove("active");
  projectType ='Client'
})


internal.addEventListener('click',()=>{
  // console.log("internal")
 
  internal.classList.add("active", "toggle-btn")
  client.classList.remove("active");
  projectType = 'Internal'

})


async function getProjectById(projectId: any) {
  const url = `${domain}/api/data/v9.2/msdyn_projects(${projectId})?$select=msdyn_subject,msdyn_projectid`;

  try {
    const response = await fetch(url, {
      method: "GET",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Accept": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Content-Type": "application/json; charset=utf-8",
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Error fetching project: ${response.status} ${response.statusText}\n${errorText}`);
    }

    const data = await response.json();
    // console.log("✅ Project Name:", data.msdyn_subject);
    return data;

  } catch (error:any) {
    // console.error("❌ Fetch error:", error.message);
  }
}


const populateProjectTaskList = (filteredKeys: { value: any[]; }) => {
  // Ensure that filteredKeys is an array
  if (Array.isArray(filteredKeys.value) && filteredKeys.value.length > 0) {
    // Clear existing options in selectTaskizeInstance

    // Add options dynamically
    filteredKeys.value.forEach((_project) => {});

    // After adding all options, refresh the dropdown
  } else {
    console.error("filteredKeys is either not an array or it's empty.");
  }
};


function populateProjectTaskListNew(tasks: any[]) {
  dropdownListTask.innerHTML = ""; // Clear previous tasks
  tasks.forEach((task) => {
    let div = document.createElement("div");
    div.textContent = task.msdyn_subject;
    div.style.fontSize = "12px";
    div.style.color = "rgb(84, 84, 84)";
    div.id = `${task.value}`;
    div.onclick = function () {
      searchInputTask.value = task.msdyn_subject;
      selectedProjectTaskName = task.msdyn_subject
      dropdownListTask.style.display = "none";

      selectedProjectTaskIdNew = task.msdyn_projecttaskid;
    };
    dropdownListTask.appendChild(div);
  });
}



function filterOptionsProject() {
  let filter = searchInputProject.value.toLowerCase();
  let items = dropdownListProject.getElementsByTagName("div");

  for (let i = 0; i < items.length; i++) {
    let txtValue = items[i].textContent || items[i].innerText;
    items[i].style.display = txtValue.toLowerCase().includes(filter) ? "" : "none";
  }
}


function filterOptionsTask() {
  let filter = searchInputTask.value.toLowerCase();
  let items = dropdownListTask.getElementsByTagName("div");

  for (let i = 0; i < items.length; i++) {
    let txtValue = items[i].textContent || items[i].innerText;
    items[i].style.display = txtValue.toLowerCase().includes(filter) ? "" : "none";
  }
}


function showDropdownProject() {
  dropdownListProject.style.display = "block";
  populateDropdown(options);
}

const fetchMatchingProjects = async (searchTerm: any) => {
  // console.log(searchTerm)
  if (searchTerm.length < 3) {
    return; // Only fetch data when at least 3 characters are typed
  }

  // console.log("Fetching projects for search term:", searchTerm);

  try {
    const projectsResponse = await fetch(
      `${domain}/api/data/v9.2/msdyn_projects?$select=msdyn_subject,msdyn_projectid&$filter=contains(msdyn_subject, '${searchTerm}')`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`, // Ensure token is set properly
        },
      }
    );

    if (!projectsResponse.ok) {
      const errorText = await projectsResponse.text();
      throw new Error(`Projects fetch failed: ${errorText}`);
    }

    const projectsData = await projectsResponse.json();
    // console.log("Filtered projects:", projectsData);
    let newOption: any[] = [];
    let projectnameArray: { [x: number]: any; }[] = [];
    projectsData.value.forEach((each: { msdyn_projectid: any; msdyn_subject: any; }) => {
      newOption.push({ value: each.msdyn_projectid, text: each.msdyn_subject });
      projectnameArray.push({ [each.msdyn_projectid]: each.msdyn_subject });
    });

    options = newOption;
    // console.log(newOption);
    populateDropdown(options);

    // console.log("Formatted Project List:", projectnameArray);

    // Populate projectInput datalist

    projectsData.value.forEach((_project: any) => {});

    // Refresh dropdown
    // selectizeInstance.refreshOptions(false);

    const selectedProjectEntry = projectnameArray.find((obj) => Object.values(obj)[0] === searchTerm);

    // console.log(selectedProjectEntry)

    if (!selectedProjectEntry) {
      console.warn("No matching project ID found for selected project.");
      return;
    }

    const selectedProjectId = Object.keys(selectedProjectEntry)[0];
    selectedProjectIdTable = selectedProjectId;

    // console.log("Selected Project ID:", selectedProjectId);

    fetchProjectTasks(selectedProjectId);
  } catch (error) {
    console.error("Error fetching projects:", error);
  }
};

searchInputProject.addEventListener("keyup", filterOptionsProject); 
searchInputProject.addEventListener("click", showDropdownProject);
searchInputProject.addEventListener("focus", ()=>{
  dropdownListTask.style.display = 'none'
});
searchInputTask.addEventListener("keyup", filterOptionsTask);
searchInputTask.addEventListener("click", () => {
  dropdownListTask.style.display = "block";
}); // Show task dropdown when clicked

searchInputProject.addEventListener("blur", () => {
  projectError.textContent = "";
});

searchInputTask.addEventListener("blur", () => {
  taskError.textContent = "";
});

searchInputTask.addEventListener("focus", ()=>{
  dropdownListProject.style.display = 'none'
});

searchInputProject.addEventListener("keyup", (event: any) => {
  const searchTerm = event.target.value;
  fetchMatchingProjects(searchTerm);
  if (searchTerm === ''){
    fetchProject(token)
  }
});



const fetchProjectTasks = async (selectedProjectId: any) => {
  // console.log("Project Task fetching")
  projectTaskArr = [];
  // console.log(selectedProjectId,"fetching project task")
  try {
    const responseProjectTask = await fetch(
      `${domain}/api/data/v9.0/msdyn_projecttasks?$select=msdyn_subject,_msdyn_project_value&$filter=_msdyn_project_value eq '${selectedProjectId}'`,
      {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${token}`,
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
          Prefer: "odata.include-annotations=OData.Community.Display.V1.FormattedValue",
        },
      }
    )
      .then((r) => r.json())
      .then((result) => {
        result.value.map((each: { msdyn_subject: any; }) => projectTaskArr.push(each.msdyn_subject));
        // console.log(result,"Project Task value")
      });

    let response:any = await fetch(
      `${domain}/api/data/v9.0/msdyn_projecttasks?$select=msdyn_subject,_msdyn_project_value&$filter=_msdyn_project_value eq '${selectedProjectId}'`,
      {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${token}`,
          "OData-MaxVersion": "4.0",
          "OData-Version": "4.0",
          Prefer: "odata.include-annotations=OData.Community.Display.V1.FormattedValue",
        },
      }
    )
    
    // console.log(response)
    if (response.ok){
      response =await response.json()
      

      populateProjectTaskListNew(response.value);


    }else{
      taskError.textContent = "Permission denied to access the Project Task Table."
    }
    // .then((result=>{
    //   result.value.map(each=>projectTaskArr.push(each.msdyn_subject))
    //   console.log(result)

    // }));

    //  console.log(projectTaskArr)

    // // Populate projectTaskList

    populateProjectTaskList(response);
    populateProjectTaskListNew(response.value);
  } catch (error) {
    console.error("Failed to fetch project tasks:", error);
  }
};





function populateDropdown(options: any[]) {
  // console.log(options)
  dropdownListProject.innerHTML = "";
  options.forEach((option) => {
    let div = document.createElement("div");
    div.textContent = option.text;
    div.id = `${option.value}`;
    div.style.fontSize = "12px";
    div.style.color = "rgb(84, 84, 84)";
    div.onclick = function () {
      getProjectById(option.value)
      searchInputProject.value = option.text;
      selectedProjectName = option.text
      dropdownListProject.style.display = "none";
      // dropdownListTask.style.display = "none"
      fetchProjectTasks(option.value); // Fetch tasks for selected project
      selectedProjectIdNew = option.value;
      searchInputTask.value = "";
    };
    dropdownListProject.appendChild(div);
    

  });
}



async function fetchProject(accessToken:any) {

  // console.log("project fetching called")

  try {
    
    const projectsResponse = await fetch(
      `${domain}/api/data/v9.2/msdyn_projects?$select=msdyn_subject,msdyn_projectid&$top=20`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );
    let newOption:any = [];
    if (projectsResponse.ok) {
      const projectsData = await projectsResponse.json();
      // console.log("Projects data retrieved successfully:", projectsData);

      projectsData.value.forEach((each:any) => {
        newOption.push({ value: each.msdyn_projectid, text: each.msdyn_subject });
        projectnameArray.push({ [each.msdyn_projectid]: each.msdyn_subject });
      });
      options = newOption;
      // console.log(projectnameArray); // Output: [{ id1: "Project 1" }, { id2: "Project 2" }, ...]
      populateDropdown(options);
      // selectizeInstance.clearOptions();
      projectsData.value.forEach((project: { msdyn_projectid: any; msdyn_subject: any; }) => {
        let optionObj = { value: project.msdyn_projectid, text: project.msdyn_subject };
        // console.log(optionObj)
      });

      // Refresh dropdown

      if (projectsData.value.length > 0) {
        // console.log(`Number of projects retrieved: ${projectsData.value.length}`);
      }
    } else {
      const errorText = await projectsResponse.text();
      throw new Error(`Projects fetch failed: ${errorText}`);
    }
  } catch (error) {
    console.error("Dynamics CRM API call failed:", error);
  }
  
}



async function fetchOptions(token: any) {
  // console.log("function called");
  try {
    const response = await fetch(
      `${domain}/api/data/v9.0/EntityDefinitions(LogicalName='msdyn_timeentry')/Attributes(LogicalName='hollis_projecttype')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          "OData-Version": "4.0",
          "OData-MaxVersion": "4.0",
        },
      }
    );

    const data = await response.json();

    if (data && data.OptionSet && data.OptionSet.Options) {
      const optionsArray :any= []; // Initialize an array to hold the objects
      const options = data.OptionSet.Options;

      options.forEach(function (option:any) {
        const optionObject = {
          label: option.Label.UserLocalizedLabel.Label,
          value: option.Value,
        };
        optionsArray.push(optionObject); // Append the object to the array
      });

      // console.log(optionsArray); // Output the array
      choiceOptions = optionsArray;
    } else {
      console.error("OptionSet or Options not found in response.");
    }
  } catch (error) {
    console.error("Error fetching or processing data:", error);
  }
}

let host:any ;

  const tokenRequestUserProfile = {
    scopes: ["User.Read"],
  };


   



Office.onReady(async (info) => {
  host = info.host
  switch (info.host) {
    case Office.HostType.Excel:
    case Office.HostType.PowerPoint:
    case Office.HostType.Word:
    case Office.HostType.Outlook:
      // await getFileNames()
      // await getUserData()
    
    // console.log("entered in host:",info.host)
    await accountManager.initialize()

  //    const userAccount = await accountManager.ssoGetUserIdentity(["user.read"]);
  //   const idTokenClaims = userAccount.idTokenClaims as { name?: string; preferred_username?: string };
  //     console.log(userAccount)
  //   console.log(userAccount.accessToken);


  //    const response = await fetch(`https://graph.microsoft.com/v1.0/me`, {
  //   headers: { Authorization: userAccount.accessToken },
  // });

  // if (response.ok) {
  //   // Get the user name from response JSON.
  //   const data = await response.json();
  //   const name = data.displayName;

  //   if (name) {
  //     console.log("You are now signed in as " + name + ".");
  //   }
  // } else {
  //   const errorText = await response.text();
  //   console.log("Microsoft Graph call failed - error text: " + errorText);
  // }


    const accessToken2 = await accountManager.ssoGetToken([`${domain}/user_impersonation`]); //Hollis scope
    // console.log("Access token2: ", accessToken2);
    token = accessToken2

   await fetchOptions(accessToken2);
   await fetchProject(accessToken2)


   let item:any;

   if (info.host===  Office.HostType.Outlook){
    item = Office.context.mailbox.item;
    
   }


   
   if (item && item.itemType === Office.MailboxEnums.ItemType.Appointment) {
    setEventDetails();
  }

      
      break;
  }



  // Function to fetch event details and assign them to the input fields
function setEventDetails() {
  const item :any= Office.context.mailbox.item;

  // === Description/Subject ===
  const descriptionField:any = document.querySelector(".description-field");
  if (descriptionField) {
    if (typeof item.subject === "string") {
      // Read mode
      descriptionField.value = item.subject;
      descriptionField.disabled = true;
    } else if (typeof item.subject?.getAsync === "function") {
      // Edit mode
      item.subject.getAsync((result:any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          descriptionField.value = result.value;
          descriptionField.disabled = true;
        } else {
          console.error("Failed to fetch subject:", result.error);
        }
      });
    }
  }

  // === Start Date ===
  const eventDateField :any= document.querySelector(".event-date");
  if (eventDateField) {
    if (item.start instanceof Date) {
      // Read mode
      const formattedStart = item.start.toISOString().split("T")[0];
      eventDateField.value = formattedStart;
      eventDateField.disabled = true;
    } else if (typeof item.start?.getAsync === "function") {
      // Edit mode
      item.start.getAsync((result:any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
          const startTime = new Date(result.value);
          const formattedStart = startTime.toISOString().split("T")[0];
          eventDateField.value = formattedStart;
          eventDateField.disabled = true;
        } else {
          console.error("Failed to fetch start date:", result.error);
        }
      });
    }
  }

  // === Duration ===
  const eventDurationField :any = document.querySelector(".event-duration");

  const getStartEndTime = (callback:any) => {
    let startTime:any = null;
    let endTime:any = null;

    const tryCalculate = () => {
      if (startTime && endTime && callback) {
        const duration = (endTime - startTime) / (1000 * 60); // minutes
        callback(duration);
      }
    };

    const handleAsync = () => {
      item.start.getAsync((startResult: { status: Office.AsyncResultStatus; value: string | number | Date; error: any; }) => {
        if (startResult.status === Office.AsyncResultStatus.Succeeded) {
          startTime = new Date(startResult.value);

          item.end.getAsync((endResult:any) => {
            if (endResult.status === Office.AsyncResultStatus.Succeeded) {
              endTime = new Date(endResult.value);
              tryCalculate();
            } else {
              console.error("Failed to fetch end time:", endResult.error);
            }
          });
        } else {
          console.error("Failed to fetch start time:", startResult.error);
        }
      });
    };

    if (item.start instanceof Date && item.end instanceof Date) {
      // Read mode
      startTime = item.start;
      endTime = item.end;
      tryCalculate();
    } else if (typeof item.start?.getAsync === "function" && typeof item.end?.getAsync === "function") {
      // Edit mode
      handleAsync();
    } else {
      console.error("Start or End time not available.");
    }
  };

  if (eventDurationField) {
    getStartEndTime((duration:any) => {
      eventDurationField.value = duration;
      eventDurationField.disabled = true;
    });
  }
}


  const closeBtn:any = document.getElementById("closePane");
  
  if (info.host===  Office.HostType.Outlook) {
    closeBtn.addEventListener("click", ():void => {
      
      // console.log(Office.context.ui);
      Office.context.ui.closeContainer()
    });
  }else{
    closeBtn.style.display = 'none'
  }
});


function mapProjectType(choiceOptions: any[], projectType: string) {
  // Convert projectType to lowercase
  let projectTypeValue = projectType.toLowerCase();

  // Find the matching object in choiceOptions
  const matchedOption = choiceOptions.find((option) => option.label.toLowerCase() === projectTypeValue);

  // Return the corresponding value or null if not found
  // console.log(matchedOption)
  return matchedOption ? matchedOption.value : null;
}

function getFieldValues (){
  // let fieldArray=[
  //   {date:{value:date.value,error:dateError}},
  //   {project:{value:searchInputProject.value,error:projectError}},
  //   {projectTask:{value:searchInputTask.value,error:taskError}},
  //   {duration:{value:duration.value,error:durationError}},
  //   {description:{value:description.value,error:DescriptionError}}
  // ]

  let fieldArray=[
    {value:date.value,error:dateError,errorMessage:"Please Select Date"},
    {value:searchInputProject.value,error:projectError,errorMessage:"Please select project"},
    {value:searchInputTask.value,error:taskError,errorMessage:"Please select project Task"},
   {value:duration.value,error:durationError,errorMessage:"Please Fill the duration"},
    {value:description.value,error:DescriptionError,errorMessage:"Please Fill Description"}
  ]

  return fieldArray

}



function validateField (){
  // ✅ Get the current value when the event fires
  let fieldArray=getFieldValues()
  // const latestSearchValue = searchInputProject.value;
  // console.log("Latest search input value:", latestSearchValue);
  // console.log(fieldArray,"field Validation")
  let error=[]
  for (let each of fieldArray){
    // console.log(each.value)
    if (each.value === ""){
      // console.log(each.error,"Error")
      each.error.textContent = each.errorMessage;
      error.push(each.errorMessage)
    }
  }
  // console.log(error)
  // console.log(error.length)

  return error
 
 
};

// Insert Time Entry Button
document.getElementById("insertTimeEntry")!.addEventListener("click", () => {

  let error = validateField()

  if (error.length > 0){
    // console.log("Termination of submission")
    return
  }
  
  insertButton.style.pointerEvents = "none";
  insertButton.style.opacity = "0.5";
  insertError.textContent = "Please wait while data is submitting .....";
  insertError.style.color = "black"
  
  // console.log("reach after validation")
  
  createFieldValues();
});


duration.addEventListener("focus",()=>{
  durationError.textContent = ""
})
date.addEventListener("focus",()=>{
  dateError.textContent = ""
})
description.addEventListener("focus",()=>{
  DescriptionError.textContent = ""
})








async function createFieldValues() {
  let dateElement:any = document.querySelector(".event-date");
  
  // console.log("Selected Project Type:", projectType);

  let durationElement:any = document.querySelector(".event-duration");
  let descriptionElement:any = document.querySelector(".description-field");

  let dateValue = dateElement ? dateElement.value : "";

  let durationValue = durationElement ? durationElement.value : "";
  let descriptionValue = descriptionElement ? descriptionElement.value : "";

  // Log extracted values
  // console.log("Extracted Values:");
  // console.log("Date Value:", dateValue);
  // console.log("Project Type Value:", projectTypeValue);
  // console.log("Project Value:", selectedProjectIdTable);
  // console.log("Project Task Value:", selectedProjectTaskIdTable);
  // console.log("Duration Value:", durationValue);
  // console.log("Description Value:", descriptionValue);

  // Convert date to ISO 8601 format
  let formattedDate = dateValue ? new Date(dateValue).toISOString() : null;
  // console.log(choiceOptions);

  const result = mapProjectType(choiceOptions, projectType);
  // console.log(result);
  // console.log(selectedProjectName,"selectedProjectName")
  // console.log(selectedProjectTaskName,"selectedProjectTaskName")

  const newEntryPayload = {
    hollis_projecttype: result, // Project Type
    // "msdyn_project@odata.bind": `msdyn_projects(${selectedProjectIdTable})`,
    // "msdyn_projectTask@odata.bind": `msdyn_projecttasks(${selectedProjectTaskIdTable})`,

    // "msdyn_project@odata.bind": `msdyn_projects(${selectedProjectIdNew})`,
    // "msdyn_projectTask@odata.bind": `msdyn_projecttasks(${selectedProjectTaskIdNew})`,
    "msdyn_project@odata.bind": `/msdyn_projects(${selectedProjectIdNew})`,
  "msdyn_projectTask@odata.bind": `/msdyn_projecttasks(${selectedProjectTaskIdNew})`,
  

    msdyn_date: formattedDate, // Date formatted as ISO 8601
    msdyn_duration: parseInt(durationValue, 10) || 0, // Convert duration to integer
    msdyn_description: descriptionValue, // Description
  };

  // Log the payload before posting
  // console.log("Payload to be sent:", newEntryPayload);
  // console.log(newEntryPayload);
  try {
    const response = await fetch(
      // `https://hollis-projectops-dev-01.api.crm4.dynamics.com/api/data/v9.2/msdyn_timeentries`,
      `${domain}/api/data/v9.2/msdyn_timeentries`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          "OData-Version": "4.0",
          "OData-MaxVersion": "4.0",
          "MSCRM.SuppressDuplicateDetection": "false",
        },
        body: JSON.stringify(newEntryPayload),
      }
    );

    // Office.context.ui.closeContainer();
    // console.log(response);

    

    if (!response.ok) {
      const errorText = await response.text(); // Capture response even if it's not JSON

      let error = JSON.parse(errorText);
      // console.log(error);
      insertError.style.color = "red";
      insertError.style.fontSize = "10px";
      insertError.textContent = "Error Encountered while submitting the data please insert proper data.";

      throw new Error(`Error creating record: ${response.status} ${response.statusText} - ${errorText}`);
    } else {
      insertError.textContent = "Time entry inserted successfully. Please close the task pane using the cross icon at the top right corner.";
      insertError.style.color = "Green";
      insertError.style.marginBottom = "10px";
      const insertButton = document.getElementById("insertTimeEntry") as HTMLElement;
insertButton.style.pointerEvents = "none";
insertButton.style.opacity = "0.5";

     
setTimeout(() => {
  if (host === Office.HostType.Outlook)
  Office.context.ui.closeContainer();
}, 3000);
    }

    

    // Check if response has content before parsing JSON
  } catch (error) {
    // console.error("Error:", error);
  }
}




