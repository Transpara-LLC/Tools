This script will pack a VKPI instance (System Files, SQL DB and also generate a restore script)

Considerations:
  1- When selecting the target folder for backing up and restoring, please use a folder where SQL has access.
  example:
    C:\Backup -- VALID
    C:\User\SpecificUser\Desktop -- INVALID
  2- If the target IIS site and source IIS site have different port bindings, a manual edit to interfaces.dbo through SQL Manager or VKPI Designer
  will be required
