# Grid configuration

The local app's **PCF Configuration** button configures the control's inputs. Applied settings persist in this browser. New local columns persist in the mock API's state file; resetting sample data removes them.

In a model-driven app, set `pageSize` (default **10**, supported range 1–250) and `gridConfigurationJson` in the PCF properties:

```json
{
  "visibleColumns": ["dms_name", "modifiedon", "dms_projectcode"],
  "editableColumns": ["dms_name", "dms_description", "dms_projectcode"],
  "filterColumns": ["dms_documenttype", "dms_documentstatus", "dms_projectcode"],
  "searchField": "dms_name",
  "tileSize": "medium"
}
```

Column values use Dataverse logical names. Omitted visible/editable columns use the server configuration. Filters default to Document type and Status, search defaults to Name, and tiles default to medium. Tile sizes are small, medium and large. Search supports text columns; choice and Boolean filters use selectors, number filters use equality, and text/date filters match the entered text. Both list and tiles use the selected display columns. Double-click a row or tile to edit its allowed document attributes; Preview remains available from the command bar.

To add a production column, create and publish it on the Document table in Dataverse, then include its logical name in the configuration. The control discovers columns from table metadata and respects `IsValidForUpdate`. The local **Create column** feature supports text, multiline text, integer, decimal, Boolean and date columns for development. It creates local metadata, not a Dataverse schema change. Production configuration and schema discovery still require tenant verification.

Pagination requests one page at a time using API continuation links. Previous/Next replaces the displayed page; search, filters, sorting, refresh and page-size changes return to page 1. Selection clears on page navigation.

## Office preview images

`DocumentThumbnail.tsx` draws representative thumbnails with React elements and CSS. Word uses the filename, paper and placeholder text lines; Excel uses a fixed sample table; PowerPoint uses a fixed slide layout and sample chart. These images are not extracted from DOCX/XLSX/PPTX files and do not show their actual contents. Uploaded Office documents use the same representative artwork. The mock's Office file generation creates valid document packages separately and does not render thumbnails.

Photos display the downloaded image bytes. The Preview command downloads supported files for image, PDF, video and text viewing; Office documents download for opening in their applications.

Accurate Office thumbnails would require a document-rendering service or a connected storage provider's thumbnail API, then displaying the returned image. No Office rendering service or authenticated Graph thumbnail provider is connected in this app.
