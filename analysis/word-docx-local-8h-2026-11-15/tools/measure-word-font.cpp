// Diagnostic only. Measure the repository's licensed fallback faces using the
// documented Windows APIs; no Microsoft embedded font or product asset is used.
#include <windows.h>
#include <dwrite_1.h>
#include <wrl/client.h>
#include <iostream>
#include <iomanip>
#include <string>
#include <vector>
#include <stdexcept>
using Microsoft::WRL::ComPtr;
void check(HRESULT hr) { if (FAILED(hr)) throw std::runtime_error("DirectWrite HRESULT failure: " + std::to_string(static_cast<unsigned long>(hr))); }
int wmain(int argc, wchar_t** argv) {
  try {
    const std::wstring root = argc > 1 ? argv[1] : L"src/assets/fonts";
    ComPtr<IDWriteFactory> factory;
    check(DWriteCreateFactory(DWRITE_FACTORY_TYPE_ISOLATED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(factory.GetAddressOf())));
    const wchar_t* files[] = {L"texgyreadventor-regular.otf",L"texgyreadventor-bold.otf",L"texgyreadventor-italic.otf",L"texgyreadventor-bolditalic.otf"};
    const char* styles[] = {"regular","bold","italic","bolditalic"};
    std::vector<UINT32> codes; for (UINT32 c=32;c<127;c++) codes.push_back(c);
    std::cout << std::setprecision(10) << "{\"method\":\"Native DirectWrite design versus GDI compatible advances; licensed fallback only; diagnostic, not product coefficients\",\"fonts\":[";
    for (int f=0;f<4;f++) {
      ComPtr<IDWriteFontFile> file; const auto path=root+L"/"+files[f];
      check(factory->CreateFontFileReference(path.c_str(),nullptr,&file));
      BOOL supported; DWRITE_FONT_FILE_TYPE fileType; DWRITE_FONT_FACE_TYPE faceType; UINT32 faceCount;
      check(file->Analyze(&supported,&fileType,&faceType,&faceCount)); if (!supported || !faceCount) throw std::runtime_error("Unsupported fallback font");
      IDWriteFontFile* raw=file.Get(); ComPtr<IDWriteFontFace> face;
      check(factory->CreateFontFace(faceType,1,&raw,0,DWRITE_FONT_SIMULATIONS_NONE,&face));
      ComPtr<IDWriteFontFace1> face1; check(face.As(&face1)); DWRITE_FONT_METRICS metrics; face->GetMetrics(&metrics);
      std::vector<UINT16> glyphs(codes.size()); check(face->GetGlyphIndices(codes.data(),static_cast<UINT32>(codes.size()),glyphs.data()));
      std::vector<INT32> design(codes.size()); check(face1->GetDesignGlyphAdvances(static_cast<UINT32>(glyphs.size()),glyphs.data(),design.data(),FALSE));
      if(f)std::cout<<',';std::cout<<"{\"style\":\""<<styles[f]<<"\",\"unitsPerEm\":"<<metrics.designUnitsPerEm<<",\"hasLegacyKerning\":"<<(face1->HasKerningPairs()?"true":"false")<<",\"measurements\":[";
      const auto lineMetrics = [](const DWRITE_FONT_METRICS& m) {
        std::cout << "{\"unitsPerEm\":" << m.designUnitsPerEm << ",\"ascent\":" << m.ascent
          << ",\"descent\":" << m.descent << ",\"lineGap\":" << m.lineGap << '}';
      };
      bool first=true;
      for(float points:{9.f,10.f,11.f,12.f,14.f,16.f,18.f,28.f}) for(float dpi:{96.f,120.f,192.f}) for(int mode=0;mode<3;mode++) {
        std::vector<INT32> advances=design;
        DWRITE_FONT_METRICS line = metrics;
        if(mode)check(face->GetGdiCompatibleMetrics(points*96.f/72.f,dpi/96.f,nullptr,&line));
        if(mode)check(face1->GetGdiCompatibleGlyphAdvances(points*96.f/72.f,dpi/96.f,nullptr,mode==2,FALSE,static_cast<UINT32>(glyphs.size()),glyphs.data(),advances.data()));
        if(!first)std::cout<<',';first=false;std::cout<<"{\"points\":"<<points<<",\"dpi\":"<<dpi<<",\"mode\":\""<<(mode==0?"design":mode==1?"gdiClassic":"gdiNatural")<<"\",\"lineMetrics\":";lineMetrics(line);std::cout<<",\"advances\":[";
        for(size_t i=0;i<advances.size();i++){if(i)std::cout<<',';std::cout<<"{\"code\":"<<codes[i]<<",\"glyph\":"<<glyphs[i]<<",\"designUnits\":"<<advances[i]<<",\"points\":"<<double(advances[i])*points/metrics.designUnitsPerEm<<'}';}std::cout<<"]}";
      }
      std::cout<<"]}";
    }
    std::cout<<"]}\n";return 0;
  } catch(const std::exception& error) { std::cerr<<error.what()<<'\n';return 1; }
}
