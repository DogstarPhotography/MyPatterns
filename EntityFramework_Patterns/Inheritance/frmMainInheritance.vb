Public Class frmMainInheritance

	Private Sub frmMain_Load(sender As Object, e As System.EventArgs) Handles Me.Load

		Using context As New InheritedModelContainer
			dgvData.DataSource = context.TPTDataItems
		End Using

		Using context As New TPHModelContainer
			dgvTPH.DataSource = context.TPHDataItems
		End Using

	End Sub

	Private Sub btnAddNew_Click(sender As System.Object, e As System.EventArgs) Handles btnAddNew.Click
		'Get a context
		Using context As New InheritedModelContainer
			'Create a new PrimaryData item and set it's fields
			Dim newPrimary As New TPTPrimaryData
			newPrimary.Name = txtName.Text & "_Primary"
			newPrimary.Value = txtValue.Text
			newPrimary.Identifier = txtIdentifier.Text
			'Create a new SecondaryData item, set it's fields, and add a refernce to it to the new PrimaryData item
			Dim newSecondary As New TPTSecondaryData
			newSecondary.Name = newPrimary.Name & "_Secondary"
			newSecondary.Value = newPrimary.Value
			newSecondary.Identifier = newPrimary.Identifier
			newPrimary.SecondaryDatas.Add(newSecondary)
			'Now add the new PrimaryData item to the base DataItems table and then save changes
			'	There is no PrimaryDatas table because it inherits from DataItem and can thus be stored in the DataItems table
			'	with any PrimaryData specific fields being stored in the DataItems_PrimaryData table and linked to the DataItems table
			context.TPTDataItems.AddObject(newPrimary)
			context.SaveChanges()
			dgvData.DataSource = context.TPTDataItems
		End Using
	End Sub

	Private Sub btnNewTPHItem_Click(sender As System.Object, e As System.EventArgs) Handles btnNewTPHItem.Click
		'Get a context
		Using context As New TPHModelContainer
			'Create a new PrimaryData item and set it's fields
			Dim newPrimary As New TPHPrimaryData
			newPrimary.Name = txtTPHName.Text & "_Primary"
			newPrimary.Value = txtTPHValue.Text
			newPrimary.PrimaryIdentifier = txtTPHIdentifier.Text
			'Create a new SecondaryData item, set it's fields, and add a refernce to it to the new PrimaryData item
			Dim newSecondary As New TPHSecondaryData
			newSecondary.Name = newPrimary.Name & "_Secondary"
			newSecondary.Value = newPrimary.Value
			newSecondary.SecondaryIdentifier = newPrimary.PrimaryIdentifier
			newPrimary.TPHSecondaryDatas.Add(newSecondary)
			'Now add the new PrimaryData item to the base DataItems table and then save changes
			'	There is no PrimaryDatas table because it inherits from DataItem and can thus be stored in the DataItems table
			'	with any PrimaryData specific fields being stored in the DataItems_PrimaryData table and linked to the DataItems table
			context.TPHDataItems.AddObject(newPrimary)
			context.SaveChanges()
			dgvTPH.DataSource = context.TPHDataItems
		End Using
	End Sub
End Class
