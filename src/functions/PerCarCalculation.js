
module.exports =  (qualifiedRM, formData) => {

 
    qualifiedRM.forEach((record) => {

        let soldCar = parseInt(record["Grand Total"]);
        let perCarIncentive = 0;
    
        // Find the appropriate incentive based on the exact number of cars sold
        formData.carIncentive.forEach((incentive) => {
    
          if (soldCar == parseInt(incentive.cars)) {
            perCarIncentive = parseInt(incentive.incentive);
            // console.log("perCarIncentive ::::::::::::::::::",perCarIncentive);
          }
        });
    
    
        const lastIncentive = formData.carIncentive[formData.carIncentive.length - 1].incentive;
          if (soldCar > parseInt(formData.carIncentive[formData.carIncentive.length - 1].cars)) {
              perCarIncentive = lastIncentive;
          }
    
        // Add the incentive to the record
        record["Per Car Incentive"] = perCarIncentive;
        record["Total PerCar Incentive"] = soldCar * perCarIncentive;
        record["Total Incentive"] = soldCar * perCarIncentive;
      });
    
    

    
    return qualifiedRM;
    }
    
