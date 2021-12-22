#include <shlwapi.h>
#include <thumbcache.h> // For IThumbnailProvider.
#include <wrl/client.h> // For ComPtr
#include <new>
#include <atomic>
#include <vector>

template <typename T>
using ComPtr = Microsoft::WRL::ComPtr<T>;

// this thumbnail provider implements IInitializeWithStream to enable being hosted
// in an isolated process for robustness
class CQOIThumbProvider final : public IInitializeWithStream, public IThumbnailProvider {
    __pragma(pack(push, 1)) struct QOIRgba {
        uint8_t r, g, b, a;

        inline size_t Hash() const {
            return static_cast<size_t>(r) * 3 +
                   static_cast<size_t>(g) * 5 +
                   static_cast<size_t>(b) * 7 +
                   static_cast<size_t>(a) * 11;
        }
    } __pragma(pack(pop));

    struct QOIImage {
        static const int ColorspaceSRGB     = 0;
        static const int ColorspaceLinear   = 1;

        uint32_t             width;
        uint32_t             height;
        uint32_t             channels;
        uint32_t             colorSpace;
        std::vector<QOIRgba> pixels;
    };

public:
    CQOIThumbProvider()
        : mReferences(1)
        , mStream{} {
    }

    virtual ~CQOIThumbProvider() {
    }

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) {
        static const QITAB qit[] = {
            QITABENT(CQOIThumbProvider, IInitializeWithStream),
            QITABENT(CQOIThumbProvider, IThumbnailProvider),
            { 0 },
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef() {
        return ++mReferences;
    }

    IFACEMETHODIMP_(ULONG) Release() {
        const long refs = --mReferences;
        if (!refs) {
            delete this;
        }
        return refs;
    }

    // IInitializeWithStream
    IFACEMETHODIMP Initialize(IStream* pStream, DWORD /*grfMode*/) {
        HRESULT hr = E_UNEXPECTED;  // can only be inited once
        if (!mStream) {
            // take a reference to the stream if we have not been inited yet
            hr = pStream->QueryInterface(mStream.ReleaseAndGetAddressOf());
        }
        return hr;
    }

    // IThumbnailProvider
    IFACEMETHODIMP GetThumbnail(UINT /*cx*/, HBITMAP* phbmp, WTS_ALPHATYPE* pdwAlpha) {
        QOIImage image;
        if (!this->LoadQOIImageFromStream(image)) {
            return E_FAIL;
        }

        if (!this->QOIImageToHBITMAP(image, phbmp)) {
            return E_OUTOFMEMORY;
        }

        *pdwAlpha = image.channels == 3 ? WTSAT_RGB : WTSAT_ARGB;

        return S_OK;
    }

private:
    uint8_t ReadU8() {
        uint8_t result = 0;
        ULONG bytesRead = 0;
        HRESULT hr = mStream->Read(&result, 1, &bytesRead);
        return (FAILED(hr) || bytesRead < 1) ? 0 : result;
    }

    uint32_t ReadU32_BE() {
        uint32_t a = this->ReadU8();
        uint32_t b = this->ReadU8();
        uint32_t c = this->ReadU8();
        uint32_t d = this->ReadU8();
        return a << 24 | b << 16 | c << 8 | d;
    }

    bool LoadQOIImageFromStream(QOIImage& image) {
        constexpr uint8_t   QOI_OP_INDEX    = 0x00; /* 00xxxxxx */
        constexpr uint8_t   QOI_OP_DIFF     = 0x40; /* 01xxxxxx */
        constexpr uint8_t   QOI_OP_LUMA     = 0x80; /* 10xxxxxx */
        constexpr uint8_t   QOI_OP_RUN      = 0xC0; /* 11xxxxxx */
        constexpr uint8_t   QOI_OP_RGB      = 0xFE; /* 11111110 */
        constexpr uint8_t   QOI_OP_RGBA     = 0xFF; /* 11111111 */

        constexpr uint8_t   QOI_MASK_2      = 0xC0; /* 11000000 */

        constexpr uint32_t  QOI_MAGIC       = ((uint32_t('q') << 24) | (uint32_t('o') << 16) | (uint32_t('i') <<  8) | uint32_t('f'));
        constexpr uint32_t  QOI_HEADER_SIZE = 14u;
        constexpr uint32_t  QOI_PADDING     = 8u;
        constexpr uint32_t  QOI_PIXELS_MAX  = 400000000u;

        if (!mStream) {
            return false;
        }

        const uint32_t magic = this->ReadU32_BE();
        if (QOI_MAGIC != magic) {
            return false;
        }

        image.width = this->ReadU32_BE();
        image.height = this->ReadU32_BE();
        image.channels = this->ReadU8();
        image.colorSpace = this->ReadU8();

        if (!image.width || !image.height ||
            image.channels < 3 || image.channels > 4 ||
            image.colorSpace > QOIImage::ColorspaceLinear ||
            image.height >= QOI_PIXELS_MAX / image.width) {
            return false;
        }

        const size_t numPixels = static_cast<size_t>(image.width) * image.height;
        image.pixels.resize(numPixels);

        QOIRgba index[64] = {};
        QOIRgba px = { 0, 0, 0, 255 };

        int run = 0;
        for (size_t i = 0; i < numPixels; ++i) {
            if (run > 0) {
                run--;
            } else {
                const uint8_t b1 = this->ReadU8();

                if (QOI_OP_RGB == b1) {
                    px.r = this->ReadU8();
                    px.g = this->ReadU8();
                    px.b = this->ReadU8();
                } else if (QOI_OP_RGBA == b1) {
                    px.r = this->ReadU8();
                    px.g = this->ReadU8();
                    px.b = this->ReadU8();
                    px.a = this->ReadU8();
                } else if (QOI_OP_INDEX == (b1 & QOI_MASK_2)) {
                    px = index[b1];
                } else if (QOI_OP_DIFF == (b1 & QOI_MASK_2)) {
                    px.r += ((b1 >> 4) & 0x03) - 2;
                    px.g += ((b1 >> 2) & 0x03) - 2;
                    px.b +=  (b1       & 0x03) - 2;
                } else if (QOI_OP_LUMA == (b1 & QOI_MASK_2)) {
                    const uint8_t b2 = this->ReadU8();
                    const int vg = (b1 & 0x3F) - 32;
                    px.r += vg - 8 + ((b2 >> 4) & 0x0F);
                    px.g += vg;
                    px.b += vg - 8 +  (b2       & 0x0F);
                } else if (QOI_OP_RUN == (b1 & QOI_MASK_2)) {
                    run = (b1 & 0x3F);
                }

                index[px.Hash() % 64] = px;
            }

            image.pixels[i] = px;
        }

        return true;
    }

    bool QOIImageToHBITMAP(const QOIImage& image, HBITMAP* phbmp) {
        BITMAPINFO bmi = {};
        bmi.bmiHeader.biSize = sizeof(bmi.bmiHeader);
        bmi.bmiHeader.biWidth = static_cast<LONG>(image.width);
        bmi.bmiHeader.biHeight = -static_cast<LONG>(image.height);
        bmi.bmiHeader.biPlanes = 1;
        bmi.bmiHeader.biBitCount = 32;
        bmi.bmiHeader.biCompression = BI_RGB;

        QOIRgba* dstPixels;
        HBITMAP hbmp = CreateDIBSection(nullptr, &bmi, DIB_RGB_COLORS, reinterpret_cast<void**>(&dstPixels), nullptr, 0);
        if (!hbmp) {
            return false;
        }

        const QOIRgba* srcPixels = image.pixels.data();
        // copy pixels over and swap RGBA -> BGRA
        const size_t numPixels = static_cast<size_t>(image.width) * image.height;
        for (size_t i = 0; i < numPixels; ++i) {
            dstPixels[i].b = srcPixels[i].r;
            dstPixels[i].g = srcPixels[i].g;
            dstPixels[i].r = srcPixels[i].b;
            dstPixels[i].a = srcPixels[i].a;
        }

        *phbmp = hbmp;

        return true;
    }

private:
    std::atomic_long    mReferences;
    ComPtr<IStream>     mStream;     // provided during initialization.
};


// 
HRESULT CQOIThumbProvider_CreateInstance(REFIID riid, void** ppv) {
    CQOIThumbProvider* provider = new (std::nothrow) CQOIThumbProvider();
    HRESULT hr = provider ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr)) {
        hr = provider->QueryInterface(riid, ppv);
        provider->Release();
    }
    return hr;
}