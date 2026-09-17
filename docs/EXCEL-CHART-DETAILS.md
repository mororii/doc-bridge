# Excel chart details

`create_chart` and `update_chart` accept chart `dataLabels` and per-series `dataLabels`, `trendline`, and `points`.

Data labels

Chart-level `dataLabels` are dispatched to every current member of `SeriesCollection`; Excel Chart itself is not treated as a `DataLabels()` owner. Series use `HasDataLabels` plus `DataLabels()`. Points use the singular `HasDataLabel` plus `DataLabel`. `show: true` enables labels only when absent; it does not call `ApplyDataLabels` for an already labelled series. Omitted `value`, `category`, `series`, and `percentage` fields preserve their native settings. `show: false` removes labels and cannot be combined with content fields in the same request.

Trendlines

Trendline has `action` (`add`, `update`, `delete`) and 1-based `index` for update/delete. Add types are `linear`, `exponential`, `logarithmic`, `polynomial`, `power`, and `movingAverage`. The native Add call receives positional `Type`, `Order`, and `Period`: values are primitive integers and omitted optional values are `Type.Missing`, never JSON nodes.

Polynomial order is an integer 2..6. Moving-average period is an integer 2..255 and strictly below the live series point count. These checks run before Add or Update mutates the trendline. Update writes only requested `Type`, `Order`, `Period`, name, equation, and R-squared properties; omitted properties are preserved. Requested trendline and point indexes are checked against their live collections before access.

Readback

Readback uses the exact native label owner shapes above. Null represents an unreadable native value or a chart-wide value mixed across series; it is not coerced to false or zero. Trendline state includes index, type, order, period, name, equation, and R-squared. Point state reader exposes fill/line RGB and singular-label state.

The shared chart-series reader exposes `points`, `dataLabels`, and `trendlines`. Readback selects the requested series index and the newly appended trendline for add. Delete checks that the live collection count falls by one; optional `expectedRemainingCount` additionally checks the requested final count. Surviving trendlines are reindexed and do not have stable IDs. Point inspection currently reads every point in each inspected series; large-chart latency remains unmeasured.
