module.exports =  (qualifiedRM, MGAdata, formData) => {

    function searchByID(data, id) {
        return data.find(item => item.ID == id);
    }   
   
qualifiedRM.forEach(element => {

 const result = searchByID(MGAdata, element["DSE ID"]);

 element["MGA"] = result["MGA/VEH"];
 let MGAforCalculation = result["TOTAL MGA SALE DDL"]

    const mgaValue = parseFloat(element["MGA"]);
    let MGAIncentive = 0; 
//    element["MGA Incentive"] = 0;
   

    for (const range of formData["MGAIncentive"]) {
     

        if (range.max === null) {
            if (mgaValue >= range.min) {
               MGAIncentive = parseFloat(range.incentive);
                break;
            }
        } else {
            if (mgaValue >= range.min && mgaValue < range.max) {
                MGAIncentive = parseFloat(range.incentive);
                break;
            }
        }
    
    }

    element["Total Incentive"] = parseFloat(element["Total Incentive"]) + parseFloat(element["EW Incentive"]) + parseFloat(element["CCP Incentive"])+ parseFloat(element["MSSF Incentive"]) + parseFloat(element["CDI Incentive"]) + parseFloat(element["Discount Incentive"]) + parseFloat(element["Exchange Incentive"])  + parseFloat(element["PerModel Incentive"]) +  parseFloat(element["Complaint Deduction"]) + parseFloat(element["MSR Incentive"])  ;

    element["MGA Incentive"] = ((parseFloat(MGAforCalculation)*parseFloat(MGAIncentive)))/100;

    element["Total Incentive"] = parseFloat(element["Total Incentive"]) + parseFloat(element["MGA Incentive"]);

});

return qualifiedRM;
    
};



// });

// return qualifiedRM;
    
// };
