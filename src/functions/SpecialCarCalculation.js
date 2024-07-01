
module.exports =  (qualifiedRM, formData) => {

 
    qualifiedRM.forEach(element => {
    
    
    element["SpecialCar Incentive"] = 0;
    
    let specialCarIncentive = 0;
    
    for (const model in formData.SpecialCarIncentive) {
        if (element.hasOwnProperty(model) && element[model] > 0) {
             specialCarIncentive += element[model] * formData.SpecialCarIncentive[model];
        }
    }
    
    element["SpecialCar Incentive"] =   specialCarIncentive;
    
    });
    
    return qualifiedRM;
    }
    