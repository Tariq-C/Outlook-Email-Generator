##########################################################################################################
# Created by: Tariq Chatur, Brianna Goulet
# Last Modified: October 1st, 2020
#
# Notes:
# There is no reason to edit anything in this Script
# If you wish to make adjustments. Make a copy of the script and do the work There
#

##########################################################################################################

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

###############################		FUNCTIONS		############################################################

# Function Name: email_template_form
# Input(s): None
# Output(s): Ticket Number, number of email templates
# Short Description: This Function Creates the UI for this Application
Function email_template_form ($files, $count){
	# Step 1: Definte the form
	#___________________________________________________________________________________________________________________________

	#
	#   Build the form that the user will use to genereate emails
	#

  # An array of radio buttons. One for each email template
  $radioButtons = [System.Object[]]::new($count)

  # Initialisation of the button sizes so they remain constant
  $bw = 250
  $bh = 30
  $bi = 10

  # A counter to keep track of current button
  $i = 0

  # return value which will indicate which button was selected
  $x = -1

  # Start of the form Creation, Initialisation of the form
	$Form = New-Object System.Windows.Forms.Form
	$Form.Text = 'Choose an Email Template'
	$Form.Width = 800;
	$Form.Height = 150 + ($bi + $bh)*$count;

	# Set the font of the text to be used within the form
	$Font = New-Object System.Drawing.Font("Calibri",11.5)
	$Form.Font = $Font

	# Create a text label for the Input Textbox

	$labelAreaCode = New-Object System.Windows.Forms.Label
	$labelAreaCode.Left = 40;
	$labelAreaCode.Width = 200;
	$labelAreaCode.Height = 30;
	$labelAreaCode.Top = 25;
	$labelAreaCode.Text = 'Incident/Task Number:';
	$Form.Controls.Add($labelAreaCode) # add the label to the form

	# Define a text input box for the incident number

	$textboxAreaCode = New-Object System.Windows.Forms.TextBox
	$textboxAreaCode.Left = 250;
	$textboxAreaCode.Top = 25;
	$textboxAreaCode.Height = 50;
	$textboxAreaCode.Width = 160;
	$Form.Controls.Add($textboxAreaCode) # add the input box to the form

	# Create a text label for the incident number example

	$labelAreaCodeExample = New-Object System.Windows.Forms.Label
	$labelAreaCodeExample.Left = 40;
	$labelAreaCodeExample.Width = 200;
	$labelAreaCodeExample.Height = 80;
	$labelAreaCodeExample.Top = 65;
	$labelAreaCodeExample.Text = 'Example: INC0000000 / SCTASK0000000';
	$Form.Controls.Add($labelAreaCodeExample) # add the label to the form

  # A loop that creates a radio button and labels them based on the values found in the html files
  foreach ($button in $radioButtons){
    $name = (Get-Content $files[$i] -Head 5)[-1]
    $radioButtons[$i] = New-Object System.Windows.Forms.RadioButton
    $radioButtons[$i].Left = $Form.Width - $bi - $bw;
    $radioButtons[$i].Top = $bi * ($i+1) + $bh * $i;
    $radioButtons[$i].Height = $bh;
    $radioButtons[$i].Width = $bw;
    $radioButtons[$i].Text = $name;
    $radioButtons[$i].Checked = $false
    $Form.Controls.Add($radioButtons[$i])
    $i = $i + 1
  }

  # Create the OK button
  $OKButton = new-object System.Windows.Forms.Button
  $OKButton.Left = $bi + $bw + $bi;
  $OKButton.Top = $Form.Height - $bi*2 - $bh*2;
  $OKButton.Height = $bh;
  $OKButton.Width = $bw;
  $OKButton.Text = 'OK'
  $OKButton.DialogResult=[System.Windows.Forms.DialogResult]::OK
  $Form.Controls.Add($OKButton)


	# Create a Cancel button control
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Left = $bi;
	$CancelButton.Top = $Form.Height - $bi*2 - $bh*2;
	$CancelButton.Height = $bh;
	$CancelButton.Width = $bw;
	$CancelButton.Text = 'Cancel'
	$CancelButton.DialogResult = 'Cancel'
	$Form.Controls.Add($CancelButton) # add the Cancel control to the form

	$Form.Topmost = $True                    # set the form to the foreground

	# Step 2: Open the form
	#___________________________________________________________________________________________________________________________

	$Form.Add_Shown({$Form.Activate() })     # activate / display the form

	# Step 3: Form Logic
	#___________________________________________________________________________________________________________________________

  # If Cancel is clicked then exit the form
	if('Cancel' -eq $Form.ShowDialog()){
	    # pause
	    Exit
	}else{

			# Get the ticket number and remove any dashes or spaces
			$INC = $textboxAreaCode.Text
			$INC = $INC -replace "-", ""
			$INC = $INC -replace " ", ""


      # Loop through the buttons and see which one is checked then save it to be x
      $j =  0
      foreach($button in $radioButtons){
        if($radioButtons[$j].Checked){
          $x = $j
        }
        $j++
      }

			# return the ticket number and the email selection
			return $INC, $x
    }
}

# Function Name: Generate_Email
# Input(s): String Signature, $String Incident number, Body - HTMl File
# Output(s): None
# Short Description: Template for other emails
Function Generate_Email($Signature, $INC, $Body, $header, $to){

  	# Don't touch this stuff
  	$olFolderDrafts = 16
  	$ol = New-Object -comObject Outlook.Application
  	$ns = $ol.GetNameSpace("MAPI")

  	# Creates an email Draft - Don't touch this
  	$mail = $ol.CreateItem(0)

    # Sets the to section to whatever the value in the template is
    $Mail.To = $to

  	# Sets the header to whatever the value in the template is
  	$Mail.Subject = "$INC - $header"


    # Sets the body to be whatever is in the template with the autogenerated parts
  	$Mail.HTMLBody = "<p>Good Morning/Afternoon</p><p>I'm writing in regards to Support Request $INC.</p>"

    $Mail.HTMLBody = $Mail.HTMLBody + $Body

    $Mail.HTMLBody = $Mail.HTMLBody + "<p>Thank you,</p>"

  	# Write the signature line by line
     foreach ($i in $Signature){
       $Newline = "<br>$i</br>"
       $Mail.HTMLBody = $Mail.HTMLBody + $newline
     }
  	# Saves the Email Draft - Don't touch this
  	$Mail.save()
}

# Function Name: get_sig
# Input(s): None
# Output(s): None
# Short Description: Gets the signature from the account running the script
Function get_sig{

  # Gets Signature (the .txt version)

  $signLocation = "$env:userprofile\AppData\Roaming\Microsoft\Signatures"
  $sigEnd = "*.txt"

  $Signatures = Get-ChildItem $signLocation -Filter $sigEnd
  $path = $signLocation + "\" + $Signatures[0]


  # Sets the Signature

    $Signature = Get-Content -Path $path

    return $Signature
}


# Function Name: main
# Input(s): None
# Output(s): None
# Short Description: This is the main loop
Function main {

  # Card coded values of the information stored in the template
  $LabelPosition = 5    #Button Label on line 5
  $headerPosition = 9   #Email header on Line 9
  $toPosition = 13      #Email to on Line 13
  $Offset = 1           #After the to section number of lines to skip



  #Saves the signature
  $Signature = get_sig

  # Initialise the incident number
  $INC = "INC0000000"

  # Run boolean
  $run = 1

  # Value to determine if an email should be made or not
  $i = -1

  # Find the scipts position in the file system
  $Source = Get-Location

  # Find the folder with the email Email_Templates
  $StrSource = $Source.tostring() + '/Email_Templates'

  # Make an array of the files except for Email_Template.html
  $files = Get-ChildItem -Path $StrSource -Recurse -Include *.html -Exclude 'Email_Template.html'

  # While Run is True
  while ($run -eq 1) {

    # Reset the value template selection to be none
    $i = -1

    # Save the values that return from the UI
    $results = email_template_form $files $files.Length
    $INC = $results[0]
    $i = $results[1]

    # If a template was selected do this
    if($i -ge 0){

      # Filter the content into pieces to be transfered to the email generator
      $content = Get-Content $files[$i]
      $temp = $content.Length - $toPosition - $Offset
      $body = Get-Content $files[$i] -Tail $temp
      $header = (Get-Content $files[$i] -Head $headerPosition)[-1]
      $to = (Get-Content $files[$i] -Head $toPosition)[-1]

      # Generate the email template
      Generate_Email $Signature $INC $body $header $to
    }

  }
}




###############################		Start of Script		############################################################

# Call the Main Function
main
