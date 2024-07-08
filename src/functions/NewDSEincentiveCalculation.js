module.exports = (newRm, formData) => {
    const newDSEPerCarIncentive = parseInt(formData.newDSEInput[0]);
    newRm.forEach(element => {
        element['Final Incentive'] = element['Grand Total'] * newDSEPerCarIncentive;
        element["Vehicle Incentive"] =  element['Final Incentive']
        // newRm.push(element['Total Incentive']);
    });

    return newRm;

};
