
module.exports =  (qualifiedRM, formData) => {

    formData["CCP"].sort((a, b) => {
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

        element["CCP Incentive"] = 0;
        let userCCP = element["CCP"];
        const noOfCarSoldCCP = element["CCPCount"];

        for (let i = 0; i < formData["CCP"].length; i++) {
            const condition = formData["CCP"][i];
            if (condition.type === 'less' && userCCP < condition.value) {
              element["CCP Incentive"] = noOfCarSoldCCP*condition.incentive;
              break;
            } else if (condition.type === 'greater' && userCCP > condition.value) {
              element["CCP Incentive"] = noOfCarSoldCCP*condition.incentive;
              break;
            } else if (condition.type === 'range' && userCCP >= condition.min && userCCP <= condition.max) {
              element["CCP Incentive"] = noOfCarSoldCCP*condition.incentive;
              break;
            }
          }
    
    
    });
    
    return qualifiedRM;
    }