// Decodes TIFF previews off the main thread. Browsers can't display TIFF themselves.

import { decodeTiff, Transform, ChannelOrder } from "image-in-browser";

const MAX_WIDTH = 1000;

self.onmessage = ({ data: { id, bytes } }) => {
  try {
    let image = decodeTiff({ data: bytes });
    if (!image) throw new Error("Couldn't read this TIFF.");
    if (image.width > MAX_WIDTH) image = Transform.copyResize({ image, width: MAX_WIDTH });
    const rgba = image.convert({ numChannels: 4 }).getBytes({ order: ChannelOrder.rgba });
    self.postMessage({ id, width: image.width, height: image.height, rgba }, [rgba.buffer]);
  } catch (error) {
    self.postMessage({ id, error: error.message });
  }
};
