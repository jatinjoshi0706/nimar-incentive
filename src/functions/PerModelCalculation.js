
module.exports =  (qualifiedRM, formData) => {

 
qualifiedRM.forEach(element => {


element["PerModel Incentive"] = 0;

let perModelIncentive = 0;

for (const model in formData.PerModelIncentive) {
    if (element.hasOwnProperty(model) && element[model] > 0) {
        perModelIncentive += element[model] * formData.PerModelIncentive[model];
    }
}

element["PerModel Incentive"] = perModelIncentive;
element["Total Incentive"] = parseFloat(element["Total Incentive"]) + parseFloat(element["PerModel Incentive"]);

});

return qualifiedRM;
}















// if(element["ALTO"]>0){
//     element["PerModel Incentive"] =  element["PerModel Incentive"] + formData.PerModelIncentive[0]["ALTO"];
//  }
// if(element["K-10"]>0){
//     element["PerModel Incentive"] =  element["PerModel Incentive"] +formData.PerModelIncentive[0]["K-10"];
// }
// if(element["CELERIO"]>0){
//     element["PerModel Incentive"] = element["PerModel Incentive"] + formData.PerModelIncentive[0]["CELERIO"];
// }
// if(element["SWIFT"]>0){
//     element["PerModel Incentive"] =  element["PerModel Incentive"] +formData.PerModelIncentive[0]["SWIFT"];
// }
// if(element["Ertiga"]>0){
//     element["PerModel Incentive"] =  element["PerModel Incentive"] +formData.PerModelIncentive[0]["Ertiga"];
// }
// if(element["DZIRE"]>0){
//     element["PerModel Incentive"] =  element["PerModel Incentive"] +formData.PerModelIncentive[0]["DZIRE"];
// }
// if(element["WagonR"]>0){
//     element["PerModel Incentive"] =  element["PerModel Incentive"] +formData.PerModelIncentive[0]["WagonR"];
// }
// if(element["S-Presso"]>0){
//     element["PerModel Incentive"] =  element["PerModel Incentive"] +formData.PerModelIncentive[0]["S-Presso"];
// }
// if(element["EECO"]>0){
//     element["PerModel Incentive"] =  element["PerModel Incentive"] + formData.PerModelIncentive[0]["EECO"];
//  }
//  if(element["BREZZA"]>0){
//     element["PerModel Incentive"] =  element["PerModel Incentive"] + formData.PerModelIncentive[0]["EECO"];
//  }