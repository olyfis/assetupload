 
//

$(function() {
	//alert();
	var $nformVal = $("#nform");
	//if ($nformVal.length) {
		//alert();
	//}
	jQuery.validator.addMethod("myFunc", function(name2) {
		
		var position = 3;
		var match = '2';

		position--;
		var r = new RegExp('^[^\s\S]{'+position+'}'+match);

		if(! name2.match(r)){
		
	      return true;
	    }
	    return false;
	  }, "Not Valid to Start Contract.");

	
  $nformVal.validate({
	  
	  
	  
	  
    rules: {
      name2 : { 
    	  myFunc:true,
        required: true,
        minlength: 3
      },

      email: {
        required: true,
        email: true
      },
 
    },
    messages : {
      name2: {
        minlength: "Name should be at least 3 characters"
      },

      email: {
        email: "The email should be in the format: abc@domain.tld"
      },
   
    }
  });
});