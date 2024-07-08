module.exports = (qualifiedRM, MGAdata, formData) => {
    function searchByID(data, id) {
        return data.find(item => item.ID == id);
    }

    qualifiedRM.forEach(element => {
        // console.log("MGAdata");
        // console.log(MGAdata);
        const result = searchByID(MGAdata, element["DSE ID"]);

        if (result) {
            element["MGA"] = result["MGA/VEH"];
            let MGAforCalculation = result["TOTAL MGA SALE DDL"];

            const mgaValue = parseFloat(element["MGA"]);
            let MGAIncentive = 0;

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

            element["MGA Incentive"] = ((parseFloat(MGAforCalculation) * parseFloat(MGAIncentive))) / 100;
            element["TOTAL MGA"] = MGAforCalculation;
            // Optionally, you can add this if you need to update the total incentive
            // element["Total Incentive"] = parseFloat(element["Total Incentive"]) + parseFloat(element["MGA Incentive"]);

        } else {
            console.error(`No matching MGA data found for DSE ID: ${element["DSE ID"]}`);
            element["MGA"] = 0;
            element["MGA Incentive"] = 0;
        }

        // console.log(`element["MGA Incentive"]`);
        // console.log(element["MGA Incentive"]);
    });

    return qualifiedRM;
};
