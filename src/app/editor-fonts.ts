/**
 * Free metric-compatible fonts for faithful on-screen editing: when Arial, Times New
 * Roman, Courier New or Calibri are not installed (Linux, Android) the browser falls
 * back to these. Faces are only downloaded when text actually uses them.
 */
import sansRegular from 'office-fonts/LiberationSans-Regular.ttf?url';
import sansBold from 'office-fonts/LiberationSans-Bold.ttf?url';
import sansItalic from 'office-fonts/LiberationSans-Italic.ttf?url';
import sansBoldItalic from 'office-fonts/LiberationSans-BoldItalic.ttf?url';
import serifRegular from 'office-fonts/LiberationSerif-Regular.ttf?url';
import serifBold from 'office-fonts/LiberationSerif-Bold.ttf?url';
import serifItalic from 'office-fonts/LiberationSerif-Italic.ttf?url';
import serifBoldItalic from 'office-fonts/LiberationSerif-BoldItalic.ttf?url';
import monoRegular from 'office-fonts/LiberationMono-Regular.ttf?url';
import monoBold from 'office-fonts/LiberationMono-Bold.ttf?url';
import calibriRegular from 'office-fonts/Carlito-Regular.ttf?url';
import calibriBold from 'office-fonts/Carlito-Bold.ttf?url';
import calibriItalic from 'office-fonts/Carlito-Italic.ttf?url';
import calibriBoldItalic from 'office-fonts/Carlito-BoldItalic.ttf?url';

const FACES: [string, string, string, string][] = [
  ['Liberation Sans', sansRegular, '400', 'normal'],
  ['Liberation Sans', sansBold, '700', 'normal'],
  ['Liberation Sans', sansItalic, '400', 'italic'],
  ['Liberation Sans', sansBoldItalic, '700', 'italic'],
  ['Liberation Serif', serifRegular, '400', 'normal'],
  ['Liberation Serif', serifBold, '700', 'normal'],
  ['Liberation Serif', serifItalic, '400', 'italic'],
  ['Liberation Serif', serifBoldItalic, '700', 'italic'],
  ['Liberation Mono', monoRegular, '400', 'normal'],
  ['Liberation Mono', monoBold, '700', 'normal'],
  ['Carlito', calibriRegular, '400', 'normal'],
  ['Carlito', calibriBold, '700', 'normal'],
  ['Carlito', calibriItalic, '400', 'italic'],
  ['Carlito', calibriBoldItalic, '700', 'italic'],
];

let registered = false;

export function registerEditorFonts(): void {
  if (registered || typeof FontFace === 'undefined') return;
  registered = true;
  for (const [family, url, weight, style] of FACES) {
    try {
      document.fonts.add(new FontFace(family, `url(${url})`, { weight, style, display: 'swap' }));
    } catch {
      /* unsupported font loading API – system fonts are used */
    }
  }
}
