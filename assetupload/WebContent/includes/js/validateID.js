 
//

$(function() {

	
	
	var $nformVal = $("#actionform");
	/***********************************************************************************************************************************/
	
	$.validator.addMethod("checkDate", function(eDate, element) {
        //var startDate = $('#date_open').val();
		var startDate = new Date();
		
        return Date.parse(startDate) <= Date.parse(eDate) || eDate == "";
    }, "* End date must be after start date");
	/***********************************************************************************************************************************/

	
	/***********************************************************************************************************************************/

	jQuery.validator.addMethod("myFunc", function(id) {
		
		var position = 3;
		var match = '2';

		position--;
		var r = new RegExp('^[^\s\S]{'+position+'}'+match);

		if(! id.match(r)){
		
	      return true;
	    }
	    return false;
	  }, "Not Valid to Start Contract.");

	/***********************************************************************************************************************************/
  $nformVal.validate({
	  
	  
	  
	  
    rules: {
      id : { 
    	  myFunc:true,
        required: true,
        minlength: 15
      },

      email: {
        required: true,
        email: true
      },
 
      eDate: {
    	  checkDate:true,
          required: true,
          
        },
      
      
    },
    messages : {
      id: {
        minlength: "Name should be at least 15 characters"
        	 
      },
      
      id: {    
          	myFunc: "ID cannot have a 2 in the third position of the contract ID."
        },
      eDate: {    
        	checkDate: "Date Enter must be later than today."
        },

      email: {
        email: "The email should be in the format: abc@domain.tld"
      },
   
    }
  });
});