﻿ALTER TABLE [dbo].[Branches]
    ADD CONSTRAINT [FK_RootBranch] FOREIGN KEY ([RootIndex]) REFERENCES [dbo].[Roots] ([Index]) ON DELETE NO ACTION ON UPDATE NO ACTION;

