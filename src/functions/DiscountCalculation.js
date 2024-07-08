
module.exports = (qualifiedRM, formData) => {
    qualifiedRM.forEach(element => {
        element["Vehicle Incentive % Slabwise"] = 0;
        element["Total Vehicle Incentive Amt. Slabwise"] = 0;
        let userValue = element["AVG. Discount"];
        // console.log("userValue");
        // console.log(userValue);
        const TotalVehicleIncentive = element["Total PerCar Incentive"] + element['SpecialCar Incentive']
        for (const incentive of formData.DiscountInputs) {
            //soon scenario
            if (incentive.max === null) {
                if (userValue >= incentive.min) {
                    element["Vehicle Incentive % Slabwise"] = incentive.incentive;
                    // console.log(element["Vehicle Incentive % Slabwise"])
                    element["Total Vehicle Incentive Amt. Slabwise"] = (TotalVehicleIncentive * incentive.incentive) / 100;
                    break;
                }
            } else {
                if (userValue >= incentive.min && userValue < incentive.max) {
                    element["Vehicle Incentive % Slabwise"] = incentive.incentive;
                    element["Total Vehicle Incentive Amt. Slabwise"] = (TotalVehicleIncentive * incentive.incentive) / 100;
                    break;
                }
            }
        }
        // console.log(`element["Discount Incentive"]`);
        // console.log(element["Discount Incentive"]);

    });
    return qualifiedRM;
}