# Excel picture editing

`update_picture` supports `crop` `{left,top,right,bottom}` in PictureFormat points and `path` replacement from an absolute supported local image file. Crop values are relative to original image geometry, not current displayed size.

Replacement creates/configures a temporary new native picture first, carries logical name, placement, size, rotation, lock-aspect setting, and compatible crop values, hands off its logical name, then deletes the old picture. If new-picture creation/preparation fails, the old picture remains. Native shape identity and z-order are not preserved. The source filename is application evidence only; no byte-equivalence claim is made. Inspection reads crop and geometry when COM exposes them.

Native Excel/COM execution was intentionally not run.