module.exports = (newRm, formData) => {
    // newDSEInput[0]
    console.log("newRm")
    console.log(newRm)
    const newDSEPerCarIncentive = parseInt(formData.newDSEInput[0]);
    newRm.forEach(element => {
        element['Total Incentive'] = element['Grand Total'] * newDSEPerCarIncentive;
        newRm.push(element['Total Incentive']);
    });

    return newRm;

};
