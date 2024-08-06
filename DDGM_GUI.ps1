Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$exchangedd_form = New-Object System.Windows.Forms.Form
$exchangedd_form.Text ='Dynamic Distribution Group Membership Tool'
$exchangedd_form.Width = 600
$exchangedd_form.Height = 400
$exchangedd_form.AutoSize = $true

$exchangedd_Label = New-Object System.Windows.Forms.Label
$exchangedd_Label.Text = "Dynamic Distribution Group Membership"
$exchangedd_Label.Location  = New-Object System.Drawing.Point(2,10)
$exchangedd_Label.AutoSize = $true
$exchangedd_form.Controls.Add($exchangedd_Label)

$exchangedd_ComboBox = New-Object System.Windows.Forms.ComboBox
$exchangedd_ComboBox.Width = 300

$Groups = Get-DynamicDistributionGroup
Foreach ($Group in $Groups) {
    $exchangedd_ComboBox.Items.Add($Group.Name);
}
$exchangedd_ComboBox.Location  = New-Object System.Drawing.Point(2,30)
$exchangedd_form.Controls.Add($exchangedd_ComboBox)

$exchangedd_Label2 = New-Object System.Windows.Forms.Label
$exchangedd_Label2.Text = "Members:"
$exchangedd_Label2.Location  = New-Object System.Drawing.Point(2,60)
$exchangedd_Label2.AutoSize = $true
$exchangedd_form.Controls.Add($exchangedd_Label2)

$exchangedd_Label3 = New-Object System.Windows.Forms.Label
$exchangedd_Label3.Text = ""
$exchangedd_Label3.Location  = New-Object System.Drawing.Point(2,80)
$exchangedd_Label3.AutoSize = $true
$exchangedd_form.Controls.Add($exchangedd_Label3)

$exchangedd_Button = New-Object System.Windows.Forms.Button
$exchangedd_Button.Location = New-Object System.Drawing.Size(302,29)
$exchangedd_Button.Size = New-Object System.Drawing.Size(120,23)
$exchangedd_Button.Text = "Check Members"
$exchangedd_form.Controls.Add($exchangedd_Button)

$exchangedd_Button.Add_Click({
    $FTE = $FTE = Get-DynamicDistributionGroup $exchangedd_ComboBox.Text
    $MEM = Get-Recipient -RecipientPreviewFilter $FTE.RecipientFilter -OrganizationalUnit $FTE.RecipientContainer
    $number = 0
    $newgroup = $exchangedd_ComboBox.Text
    if ($OldGroup -ne $newgroup) {
        $exchangedd_Label3.Text = ""
    }
    foreach ($Member in $MEM) {
        if ($number = 0) {
            $exchangedd_Label3.Text = $Member
            $number = 1
            $OldGroup = $newgroup
        }
        else {
            $exchangedd_Label3.Text = $exchangedd_Label3.Text + $Member + [environment]::NewLine
        }
    }
})

$exchangedd_form.ShowDialog()
