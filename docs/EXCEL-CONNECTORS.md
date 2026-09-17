# Excel connectors

Public operations are `create_connector` and `update_connector`. Existing `delete_shape` deletes a connector by name.

Endpoint input

`connectorType` is `straight`, `elbow`, or `curve`. Each required creation endpoint, and either optional update endpoint, is one of:

- `{ "x": 120, "y": 80 }` for a free point in Excel points.
- `{ "shape": "Process", "site": 1 }` for a named shape and its 1-based Excel connection-site index.

A named target must exist in the same Shapes collection, cannot be the connector itself, and its requested site must be within its actual `ConnectionSiteCount`. Both requested named endpoints are validated before connector mutation. A requested creation `name` must not already name any shape on the sheet.

Native behavior

Creation passes free endpoint coordinates directly to `Shapes.AddConnector`. Attachment and detachment use `Shape.ConnectorFormat.BeginConnect`, `EndConnect`, `BeginDisconnect`, and `EndDisconnect`; these methods are not members of `Shape`. Updating a free endpoint reads the complete axis-aligned box and `HorizontalFlip`/`VerticalFlip`, retains the omitted endpoint, then writes a non-negative bounding box and changes flips when line direction requires it. Existing omitted named attachments are saved and reapplied when the other end changes. Existing line formatting continues through `ApplyShapeFormatting`.

No `RerouteConnections` call is made: caller-selected connection sites remain exact requests. Excel can adjust position and size when connecting a target; final verification reports the actual result rather than assuming the requested free point survived that native adjustment.

Inspection and verification

`connectors` inspection reports each end's `connected` value and `connectedState` (`connected`, `disconnected`, `mixed`, or `unreadable`). Mixed and unreadable flags are explicit and are never converted into `false`. For unrotated connectors, begin/end coordinates are calculated from the bounding box plus flips, rather than assuming `(left,top)` is always the beginning. Attached ends are verified by target name and site. Free ends are verified by actual readback coordinates and disconnected state.

Coordinate edits and coordinate inspection are unavailable for rotated connectors (or unreadable box/flip values), because an accurate rotated transform is not implemented. The readback records `connectorGeometryUnavailable`; non-zero rotation additionally records `connectorGeometryRotationUnsupported`, and a requested free-coordinate edit fails with a connector geometry error. Native Excel/COM execution was not run by these tests.

Existing `delete_shape` deletes connectors. Incremental generic shape restoration does not recreate connector topology and explicitly reports this limitation; the existing workbook recovery artifact is the fallback. Undoing a newly created connector can still delete that new shape.
