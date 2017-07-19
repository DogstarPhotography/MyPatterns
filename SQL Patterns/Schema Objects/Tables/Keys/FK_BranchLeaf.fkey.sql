ALTER TABLE [dbo].[Leaves]
    ADD CONSTRAINT [FK_BranchLeaf] FOREIGN KEY ([BranchIndex]) REFERENCES [dbo].[Branches] ([Index]) ON DELETE NO ACTION ON UPDATE NO ACTION;

