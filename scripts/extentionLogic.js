
/*global variables*/
let results = [];
let mainUrl = "";
const queryTitle = "My Queries/Features AutoCarRentSystemCMMI";
/**/


VSS.init({
    explicitNotifyLoaded: true,
    usePlatformScripts: true
});

let loadData = function(VSS_Service, TFS_Wit_WebApi, TFS_WitContracts,TFS_WitClient) {
	let projectId = VSS.getWebContext().project.id;
	let projectName = VSS.getWebContext().project.name;
	let colectionUrl = VSS.getWebContext().collection.uri;
	mainUrl = colectionUrl+projectName+"/_workitems/edit/";
	let witClient = VSS_Service.getCollectionClient(TFS_WitClient.WorkItemTrackingHttpClient);
	witClient.getQuery(projectId, queryTitle, TFS_WitContracts.QueryExpand.All).then((queryItm)=> {
		return witClient.queryById(queryItm.id, projectId);
    }).then((recursiveQryResult)=>{
		treeRelations = recursiveQryResult.workItemRelations;
		treeItemIds = $.map(treeRelations, function (item, index) {
			return item.target.id;
		});
		console.log(treeItemIds);
		let itemData = VSS_Service.getCollectionClient(TFS_Wit_WebApi.WorkItemTrackingHttpClient);
		itemData.getWorkItems(treeItemIds,null,new Date(),TFS_WitContracts.WorkItemExpand.All).then(
			function(workItems){
				results = workItems;
			});
	});
}

VSS.require(["VSS/Service", "TFS/WorkItemTracking/RestClient","TFS/WorkItemTracking/Contracts","TFS/WorkItemTracking/RestClient"], loadData);

VSS.notifyLoadSucceeded();
VSS.ready(function() {
    document.getElementById("name").innerText = VSS.getWebContext().user.name;
});


function createViewModel(dataTree){
	var self = this;
	self.epics = ko.observableArray(dataTree);
	var report = document.getElementById("report");
	report.style.display = "block";
}

var formatDate = function(date){
	return date.getFullYear() +"-"+(date.getMonth()+1)+"-"+date.getDate();
}

var getRequirementState = function(state){
	let translatedState = "";
	switch(state){
		case "Proposed":
			translatedState = "Pasiūlytas";
			break;
		case "Active":
			translatedState = "Aktyvus";
			break;
		case "Resolved":
			translatedState = "Atliktas";
			break;
		case "Closed":
			translatedState = "Uždarytas";
			break;
		default:
			translatedState = state;
	}
	return translatedState;
}

var formDataTree = function(data){
	let tree = [];
	data.map((item)=>{
		let featureIds = [];
		if(item.fields["System.WorkItemType"] === "Epic"){
			item.fields.Title = item.fields["System.Title"];
			item.fields.Description = item.fields["System.Description"];
			item.fields.ItemUrl = mainUrl + item.id;
			item.fields.State = getRequirementState();
			if(typeof(item.fields['Microsoft.VSTS.Common.ActivatedDate']) !== "undefined"){
				item.fields.ActivationDate = item.fields['Microsoft.VSTS.Common.ActivatedDate'] ==="NaN-NaN-NaN" ? "" : formatDate(new Date(item.fields['Microsoft.VSTS.Common.ActivatedDate']));
			}
			if(typeof(item.relations) !== "undefined" && item.relations.length !== 0){
				featureIds = item.relations.map((feature)=>{
					if(feature.attributes.name === "Child"){
						return parseInt(feature.url.split('/')[feature.url.split('/').length-1]);
					}					
				});
				data.map((it)=>{
					let requirementIds =[];
					if(featureIds.includes(it.id)){
						featureIds.splice(featureIds.indexOf(it.id),1);
						it.fields.ActivationDate = "";
						if(typeof(it.fields['Microsoft.VSTS.Common.ActivatedDate'])!== "undefined"){
							it.fields.ActivationDate = it.fields['Microsoft.VSTS.Common.ActivatedDate'] ==="NaN-NaN-NaN" ? "" : formatDate(new Date(it.fields['Microsoft.VSTS.Common.ActivatedDate'])) ;
						}
						it.fields.ItemUrl = mainUrl + it.id;
						if(typeof(item.features) === "undefined"){
							item.features =[];
						}
						item.features.push(it);
						if(typeof(it.relations) !== "undefined" && it.relations.length !== 0){
							requirementIds = it.relations.map((requirement)=>{
								if(requirement.attributes.name === "Child"){
									return parseInt(requirement.url.split('/')[requirement.url.split('/').length-1]);
								}
							});
							data.map((i)=>{
								if(requirementIds.includes(i.id)){
									requirementIds.splice(requirementIds.indexOf(i.id),1);
									i.fields.ActivationDate = "";
									if(typeof(i.fields['Microsoft.VSTS.Common.ActivatedDate']) !== "undefined"){
										i.fields.ActivationDate = i.fields['Microsoft.VSTS.Common.ActivatedDate'] ==="NaN-NaN-NaN" ? "" : formatDate(new Date(i.fields['Microsoft.VSTS.Common.ActivatedDate']));
									}
									i.fields.ItemUrl = mainUrl + i.id;
									i.fields.State = getRequirementState(i.fields['System.State']);
									if(typeof(it.requirements) === "undefined"){
										it.requirements =[];
									}
									it.requirements.push(i);
								}
							});
						}
					}
				});
			}
			tree.push(item);
		}
	});
	return tree;
};

var generateReport = function(){
	console.log("generate report button was clicked");
	if(results.length !== 0){
		var body = document.getElementById("mainBody");
		body.style.overflow = "auto";
		let dataTree = formDataTree(results);
		console.log(dataTree);
		ko.applyBindings(new createViewModel(dataTree));
	}
	else{
		alert("Please wait while data is loading");
	}
};

var exportReport = function(){
	alert("Not Implemented yet! Please copy and paste report in the Word file manually");
	console.log("Export report button was clicked");
	/*const doc = new Document({
		sections: [
			{
				properties: {},
				children: [
					new Paragraph({
						children: [
							new TextRun("Hello World"),
							new TextRun({
								text: "Foo Bar",
								bold: true,
							}),
							new TextRun({
								text: "\tGithub is the best",
								bold: true,
							}),
						],
					}),
				],
			},
		],
	});
	Packer.toBuffer(doc).then((buffer) => {
		fs.writeFileSync("My Document.docx", buffer);
	});*/
};
