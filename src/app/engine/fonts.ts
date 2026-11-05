/** URLs of the fonts embedded into generated PDFs (hashed assets, loaded on demand). */
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
import monoItalic from 'office-fonts/LiberationMono-Italic.ttf?url';
import monoBoldItalic from 'office-fonts/LiberationMono-BoldItalic.ttf?url';
import calibriRegular from 'office-fonts/Carlito-Regular.ttf?url';
import calibriBold from 'office-fonts/Carlito-Bold.ttf?url';
import calibriItalic from 'office-fonts/Carlito-Italic.ttf?url';
import calibriBoldItalic from 'office-fonts/Carlito-BoldItalic.ttf?url';
import fallbackRegular from 'dejavu-fonts-ttf/ttf/DejaVuSans.ttf?url';
import fallbackBold from 'dejavu-fonts-ttf/ttf/DejaVuSans-Bold.ttf?url';
import cjk from 'cjk-fonts/NotoSansCJKjp-Regular.otf?url';
import emoji from '@expo-google-fonts/noto-emoji/400Regular/NotoEmoji_400Regular.ttf?url';
import type { FontKey } from '../../core/pdf-fonts';

export const FONT_URLS: Record<FontKey, string> = {
  'sans-regular': sansRegular,
  'sans-bold': sansBold,
  'sans-italic': sansItalic,
  'sans-boldItalic': sansBoldItalic,
  'serif-regular': serifRegular,
  'serif-bold': serifBold,
  'serif-italic': serifItalic,
  'serif-boldItalic': serifBoldItalic,
  'mono-regular': monoRegular,
  'mono-bold': monoBold,
  'mono-italic': monoItalic,
  'mono-boldItalic': monoBoldItalic,
  'calibri-regular': calibriRegular,
  'calibri-bold': calibriBold,
  'calibri-italic': calibriItalic,
  'calibri-boldItalic': calibriBoldItalic,
  'fallback-regular': fallbackRegular,
  'fallback-bold': fallbackBold,
  cjk,
  emoji,
};
