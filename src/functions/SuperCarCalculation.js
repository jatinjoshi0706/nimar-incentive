module.exports = (qualifiedRM, MGADataSheet, salesExcelDataSheet, formData) => {
    qualifiedRM.forEach(element => {
        const userID = element["DSE ID"];
        element["Super Car Incentive"] = 0;
        salesExcelDataSheet.forEach((data) => {
            if (data.hasOwnProperty(userID)) {
                data[userID].forEach((record) => {

                    //map criteria for qualification
                    const mappedObject = {}
                    formData.superCar.superCarCriteria.forEach(key => {
                        mappedObject[key] = false;
                    });

                    if (!(formData.SpecialCarIncentive.hasOwnProperty(record["Model Name"]))) {
                        //discount check amd update
                        if (formData.superCar.superCarCriteria.includes("zerodiscountOnVehicle")) {
                            if (record["FINAL DISCOUNT"] <= 0) {
                                mappedObject.zerodiscountOnVehicle = true;
                            }
                        }
                        // MGA check and update 
                        if (formData.superCar.superCarCriteria.includes("MGASaleGT30K")) {
                            MGADataSheet.forEach((rec) => {
                                if (rec["ID"] == userID && parseFloat(rec['MGA/VEH']) >= 30000) {
                                    mappedObject.MGASaleGT30K = true;
                                }
                            });
                        }
                        //EW check amd update
                        if (formData.superCar.superCarCriteria.includes("royalPlatinum")) {
                            if (record['Extended Warranty'] > 0) {
                                mappedObject.royalPlatinum = true;
                            }
                        }
                        //insurance check and update
                        if (formData.superCar.superCarCriteria.includes("insurance")) {
                            if (record["Insurance"] > 0) {
                                mappedObject.insurance = true;
                            }
                        }

                        //ccp check and update 
                        if (formData.superCar.superCarCriteria.includes("CCP")) {
                            if (record["CCP PLUS"] > 0) {
                                mappedObject.CCP = true;
                            }
                        }
                        //AutoCard / MSR check and update
                        if (formData.superCar.superCarCriteria.includes("MSR")) {
                            if (record["Autocard"] === 'yes' || record["Autocard"] === 'YES') {
                                mappedObject.MSR = true;
                            } else {
                                // console.log(`MSR not met for record: ${JSON.stringify(record)}`);
                            }
                        }
                        let superCarStatus = true;
                        for (let i in mappedObject) {
                            if (!mappedObject[i]) {
                                superCarStatus = false;
                                break;
                            }
                        }
                        // console.log('superCarStatus')
                        // console.log(superCarStatus)
                        const superCarIncentive = parseFloat(formData.superCar.superCarIncentive);
                        if (superCarStatus) {
                            element["Super Car Incentive"] = element["Super Car Incentive"] + superCarIncentive;
                        }
                    }
                })
            }
        })

    });
    // console.log(qualifiedRM)
    return qualifiedRM;
}