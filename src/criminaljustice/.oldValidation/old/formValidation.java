/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package criminaljustice.Validation.old;

import java.awt.Color;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

/**
 *
 * @author orion116
 */
public class formValidation
{
    public static boolean validateState(String newString)
    {
        boolean status = false;
        int i = 0;  //Initializes for loop

        String states[] = {
                            "AL", "AK", "AZ", "AR", "CA",
                            "CO", "CT", "DE", "FL", "GA",
                            "HI", "ID", "IL", "IN", "IA",
                            "KS", "KY", "LA", "MA", "MI",
                            "ME", "MD", "MN", "MS", "MO",
                            "MT", "NE", "NV", "NH", "NJ",
                            "NM", "NY", "NC", "ND", "OH",
                            "OK", "OR", "PA", "RI", "SC",
                            "SD", "TN", "TX", "UT", "VT",
                            "VA", "WA", "WV", "WI", "WY"
                          };

        for (i = 0; i < states.length; i++)
        {				   //Makes sure 2 letter abbr. is an actual state
            if (newString.equalsIgnoreCase(states[i]))
            {
                status = true;
            }
        }

        return status;
    }   
    
    public static boolean validatePhoneNumber(String phoneNo) 
    {
        boolean flag = true;
		if (phoneNo.matches("\\d{10}")) 
        {   //validate phone numbers of format "1234567890"
            flag = false;
        }
		else if (phoneNo.matches("\\d{3}[-\\.\\s]\\d{3}[-\\.\\s]\\d{4}"))
        {   //validating phone number with -, . or spaces
            flag = true;
        }
		else if (phoneNo.matches("\\d{3}-\\d{3}-\\d{4}\\s(x|(ext))\\d{3,5}")) 
        {   //validating phone number with extension length from 3 to 5
            flag = false;
        }
		else if(phoneNo.matches("\\(\\d{3}\\)-\\d{3}-\\d{4}"))
        {   //validating phone number where area code is in braces ()
            flag = false;
        }
		else
        {   //return false if nothing matches the input
            flag = false;
        }
		
        return flag;
	}
    
    public static boolean validDate(String input) 
    {
        SimpleDateFormat format = new SimpleDateFormat("MM-dd-yyyy");
        boolean flag = true;
        try 
        {
            format.parse(input);
            flag = true;
        }
        catch(ParseException e)
        {
            flag = false;
        }
        return flag;
    }
    
    public static boolean validateZipcode( int Zipcode )
    {
        boolean status = false;

        if ( ( Zipcode >= 01001 )  && ( Zipcode <= 99950 ) )
            status = true;

        return status;
    }  
    
    public static void setTextField(JTextField errorInField, boolean valid)
    {
        if (valid)
        {
            errorInField.setBackground(Color.WHITE);
            errorInField.setForeground(Color.BLACK);
        } 
        else
        {
            errorInField.setBackground(Color.RED);
            errorInField.setForeground(Color.WHITE);
            errorInField.setText(errorInField.getText());
            errorInField.requestFocus();
            errorInField.selectAll();
            
            
        }
    }
    
}
