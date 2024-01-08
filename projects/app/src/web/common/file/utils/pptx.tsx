import React, { useRef, useEffect } from 'react';
import JSZip from 'jszip';
import {
  getContentTypes,
  getSlideSize,
  loadTheme,
  processSingleSlide,
  processMsgQueue,
  setThemeContent,
  genGlobalCSS
} from './ppt';

function processPPTX(data: ArrayBuffer) {
  const zip = new JSZip();
  const zipData = zip.file('data', data);
  const filesInfo = getContentTypes(zipData);
  const slideSize = getSlideSize(zipData);
  const themeContent = loadTheme(zipData);
  setThemeContent(themeContent);

  const numOfSlides = filesInfo['slides'].length;
  const slides: string[] = [];
  for (let i = 0; i < numOfSlides; i++) {
    const filename = filesInfo['slides'][i];
    const slideHtml = processSingleSlide(zip, filename, i, slideSize);
    slides.push(slideHtml);
  }
  return slides;
}

const Pptx2Html = (data: any) => {
  const slides = processPPTX(data);
  const str = genGlobalCSS();
  const refs = useRef(null);
  useEffect(() => {
    if (refs.current) {
      processMsgQueue();
    }
  }, [refs.current]);
  return (
    <div className="pg-driver-view">
      {slides.map((i) => (
        <div key={i} dangerouslySetInnerHTML={{ __html: i }}></div>
      ))}
      <style dangerouslySetInnerHTML={{ __html: str }}></style>
      <div ref={refs}></div>
    </div>
  );
};

export default Pptx2Html;
