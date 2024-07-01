
module.exports =  (qualifiedRM, formData) => {

     
    qualifiedRM.forEach(element => {

        element["Discount Incentive"] = 0;
        let userValue = element["Discount"];


          for (const incentive of formData.DiscountInputs) {
            if (incentive.max === null) {
                if (userValue >= incentive.min) {
                    element["Discount Incentive"] = (element["Total Incentive"]*incentive.incentive)/100;
                    break;
                }
            } else {
                if (userValue >= incentive.min && userValue < incentive.max) {
                    element["Discount Incentive"] = (element["Total Incentive"]*incentive.incentive)/100;
                    break;
                }
            }
        }
    
    
    });
    
    return qualifiedRM;
    }