
module.exports =  (qualifiedRM, formData) => {

     
    qualifiedRM.forEach(element => {

        element["Complaint Deduction"] = 0;
        let userComplaintNumber = element["Complaints"];


        for (let i = 0; i < formData.ComplaintInputs.length; i++) {
            if (userComplaintNumber === parseInt(formData.ComplaintInputs[i].ComplaintNumber)) {
                element["Complaint Deduction"] = formData.ComplaintInputs[i].incentive;
            }
        }
        const lastIncentive = formData.ComplaintInputs[formData.ComplaintInputs.length - 1].incentive;
        if (userComplaintNumber > parseInt(formData.ComplaintInputs[formData.ComplaintInputs.length - 1].ComplaintNumber)) {
            element["Complaint Deduction"] = lastIncentive;
        }


        
    
    
    });
    
    return qualifiedRM;
    }