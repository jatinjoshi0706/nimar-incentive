
module.exports =  (qualifiedRM, formData) => {

    formData["MSSF"].sort((a, b) => {
        if (a.type === 'greater' && b.type === 'greater') {
            return b.value - a.value; // Descending order for 'greater'
          } else if (a.type === 'less' && b.type === 'less') {
            return a.value - b.value; // Ascending order for 'less'
          } else if (a.type === 'range' && b.type === 'range') {
            return b.max - a.max; // Descending order for 'range' based on high value
          } else if (a.type === 'range') {
            return b.value - a.max;
          } else if (b.type === 'range') {
            return b.max - a.value;
          }
          // Ensuring 'greater' types come before 'less' types in case of mixed types
          return (a.type === 'greater' ? -1 : 1);
      });
  
   
    qualifiedRM.forEach(element => {

        element["MSSF Incentive"] = 0;
        let userMSSF = element["MSSF"];
        const noOfCarSoldMSSF = element["MSSFCount"];
// console.log("noOfCarSoldMSSF");
// console.log(noOfCarSoldMSSF);
        for (let i = 0; i < formData["MSSF"].length; i++) {
            const condition = formData["MSSF"][i];
            if (condition.type === 'less' && userMSSF < condition.value) {
              element["MSSF Incentive"] = noOfCarSoldMSSF*condition.incentive;
              break;
            } else if (condition.type === 'greater' && userMSSF > condition.value) {
              element["MSSF Incentive"] = noOfCarSoldMSSF*condition.incentive;
              break;
            } else if (condition.type === 'range' && userMSSF >= condition.min && userMSSF <= condition.max) {
              element["MSSF Incentive"] = noOfCarSoldMSSF*condition.incentive;
              break;
            }
          }
    
    
    });
    
    return qualifiedRM;
    }