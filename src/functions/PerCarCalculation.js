module.exports = (qualifiedRM, formData) => {
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


    //get Special car count
    let specialCarCount = 0;
    for(let key in formData.SpecialCarIncentive){
      if(record[key] > 0){
        specialCarCount += record[key];
      }
    }


    // Add the incentive to the record
    // record["Per Car Incentive"] = perCarIncentive;
    // console.log("formData.SpecialCarIncentive");
    // console.log(formData.SpecialCarIncentive);
    // 6-2 * 2500 = 10000
    record["Total PerCar Incentive"] = Math.abs(soldCar - specialCarCount) * perCarIncentive;
    // record["Total Vehicle Incentive"] = record["Total PerCar Incentive"]  + record["Total PerCar Incentive"] ;
  });
  return qualifiedRM;
}

