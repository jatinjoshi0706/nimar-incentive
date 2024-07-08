
module.exports = (qualifiedRM, CDIdata, formData) => {

    function searchByID(data, id) {
        return data.find(item => item["DSE ID"] == id);
    }

    qualifiedRM.forEach(element => {

        const result = searchByID(CDIdata, element["DSE ID"]);
        if (result) {
            element["CDI Score"] = result["CDI"];
            const cdiScore = parseFloat(element["CDI Score"]);
            let CDIIncentive = 0;
            for (const incentive of formData.CDI) {
                if (
                    (incentive.type === 'greater' && cdiScore >= incentive.cdiValue) ||
                    (incentive.type === 'less' && cdiScore <= incentive.cdiValue) ||
                    (incentive.type === 'range' && cdiScore >= incentive.cdiMin && cdiScore <= incentive.cdiMax)
                ) {
                    CDIIncentive = incentive.incentive;
                }
            }
            element["CDI Incentive"] = CDIIncentive;
        } else {
            element["CDI Score"] = "-";
            element["CDI Incentive"] = 0;
        }

    });

    return qualifiedRM;
}