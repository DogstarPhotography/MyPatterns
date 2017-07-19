Public Class frmMainModelFirst

	Private Sub frmMain_Load(sender As Object, e As System.EventArgs) Handles Me.Load

		Using context As New DataModelContainer
			dgvData.DataSource = context.Roots
			dgvSubdata.DataSource = context.Branches
			dgvTertiary.DataSource = context.Leaves
			dgvAdditional.DataSource = context.Insects
		End Using

	End Sub

	Private Sub btnAddNew_Click(sender As System.Object, e As System.EventArgs) Handles btnAddNew.Click

		'Model summary
		'
		'Root
		'+-- Branch ---> Insect
		'    +-- Leaf ---> Insect

		''This doesn't work with the standard generated model:
		'Dim newRoot As New Root
		'newRoot.Name = txtName.Text & "_r"
		'For branches = 1 To 2
		'	Dim newBranch As New Branch
		'	newBranch.Name = newRoot.Name & "_b" & branches
		'	newRoot.Branches.Add(newBranch)
		'	For insects = 1 To 1
		'		Dim newInsect As New Insect
		'		newInsect.Name = newBranch.Name & "_i" & insects
		'		newBranch.Insects.Add(newInsect)
		'	Next
		'	For leaves = 1 To 2
		'		Dim newLeaf As New Leaf
		'		newLeaf.Name = newBranch.Name & "_l" & leaves
		'		newBranch.Leaves.Add(newLeaf)
		'		For insects = 1 To 1
		'			Dim newInsect As New Insect
		'			newInsect.Name = newLeaf.Name & "_i" & insects
		'			newLeaf.Insects.Add(newInsect)
		'		Next
		'	Next
		'Next
		'context.Roots.AddObject(newRoot)

		'To ensure that a model which contains multiple associations to a single entity can be saved to the database we need to make some changes 
		'	or one of several errors will be thrown on the call to SaveChanges

		'The first option is to programmatically insert unique ids in the key fields which will then be discarded when SaveChanges is called.
		'In order for the SQL to correctly process this we must also manually edit the relationships for the tables and set 'Enforce Foreign Key Constraint' to 'No'.
		'This is because the foreign keys in the associated entity will not refer to any actual record and will thus cause an error if the relationship is enforced
		'	and the generated database defaults to enforcing the relationships.

		'The second option is to set the foreign keys to Nullable = True in the properties window
		'This also requires us to change the relationship from one to many to zero to one to many which is potentially incorrect
		'	however we do not need to programmatically set the key fields

		Using context As New DataModelContainer

			If True = False Then
				'First Option
				Dim index As Integer = 1
				Dim newRoot As New Root
				newRoot.Name = txtName.Text & "_r"
				newRoot.Index = index : index += 1
				For branches = 1 To 2
					Dim newBranch As New Branch
					newBranch.Name = newRoot.Name & "_b" & branches
					newBranch.Index += index
					newRoot.Branches.Add(newBranch)
					For insects = 1 To 1
						Dim newInsect As New Insect
						newInsect.Name = newBranch.Name & "_i" & insects
						newInsect.Index += index
						newInsect.BranchIndex = newBranch.Index
						newBranch.Insects.Add(newInsect)
					Next
					For leaves = 1 To 2
						Dim newLeaf As New Leaf
						newLeaf.Name = newBranch.Name & "_l" & leaves
						newLeaf.Index += index
						newBranch.Leaves.Add(newLeaf)
						For insects = 1 To 1
							Dim newInsect As New Insect
							newInsect.Name = newLeaf.Name & "_i" & insects
							newInsect.Index += index
							newInsect.LeafIndex = newLeaf.Index
							newLeaf.Insects.Add(newInsect)
						Next
					Next
					context.Roots.AddObject(newRoot)
				Next
			Else
				'Second option
				Dim newRoot As New Root
				newRoot.Name = txtName.Text & "_r"
				For branches = 1 To 2
					Dim newBranch As New Branch
					newBranch.Name = newRoot.Name & "_b" & branches
					newRoot.Branches.Add(newBranch)
					For insects = 1 To 1
						Dim newInsect As New Insect
						newInsect.Name = newBranch.Name & "_i" & insects
						newBranch.Insects.Add(newInsect)
					Next
					For leaves = 1 To 2
						Dim newLeaf As New Leaf
						newLeaf.Name = newBranch.Name & "_l" & leaves
						newBranch.Leaves.Add(newLeaf)
						For insects = 1 To 1
							Dim newInsect As New Insect
							newInsect.Name = newLeaf.Name & "_i" & insects
							newLeaf.Insects.Add(newInsect)
						Next
					Next
				Next
				context.Roots.AddObject(newRoot)
			End If
			context.SaveChanges()

			dgvData.DataSource = context.Roots
			dgvSubdata.DataSource = context.Branches
			dgvTertiary.DataSource = context.Leaves
			dgvAdditional.DataSource = context.Insects
		End Using
	End Sub
End Class
