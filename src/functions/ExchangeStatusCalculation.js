
module.exports = (qualifiedRM, formData) => {
    qualifiedRM.forEach(element => {
        element["Exchange Incentive"] = 0;

        if(formData.ExchangeInputs.length !== 0){

        let userExchangeNumber = element["Exchange Status"];
        for (let i = 0; i < formData.ExchangeInputs.length; i++) {
            if (userExchangeNumber === parseInt(formData.ExchangeInputs[i].ExchangeNumber)) {
                element["Exchange Incentive"] = userExchangeNumber*formData.ExchangeInputs[i].incentive;
            }
        }
        const lastIncentive = formData.ExchangeInputs[formData.ExchangeInputs.length - 1].incentive;
        if (userExchangeNumber > parseInt(formData.ExchangeInputs[formData.ExchangeInputs.length - 1].ExchangeNumber)) {
            element["Exchange Incentive"] = userExchangeNumber*lastIncentive;
        }
    }
    });
    return qualifiedRM;
}