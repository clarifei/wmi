/**
 * ⣿⣿⡇⣿⢸⣿⢸⢡⣷⡘⡄⠻⣿⡆⡷⠚⡙⠢⡀⣙⠀⣿⡇⢹⣿⣿⡀⣿⣆⣿
 * ⣿⣿⠃⣿⢸⡿⢎⠴⠚⢿⣮⣬⣦⡉⣇⠠⣷⠂⢹⣿⢠⢇⣴⢸⣿⣿⠁⣿⣿⢸
 * ⣿⡇⡀⣿⠀⠇⢸⢀⣧⠄⣿⣿⣿⣿⣿⣆⣘⣠⣾⣇⡄⣾⣿⠘⠟⠁⠀⣿⣿⣸
 * ⣿⡇⣧⠹⡇⠀⣿⣇⠘⢀⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣇⢹⣿⡀⡖⠀⢸⡿⢃⡘
 * ⣿⣷⠹⠆⠳⠀⣿⣿⣿⣿⣿⣿⣿⠛⠻⣿⣿⣿⣿⣿⡿⠀⣿⡇⠀⠀⠘⠁⣾⣇
 * ⣿⣿⢢⡀⠱⣆⠹⣿⣿⣿⣿⣿⣿⣘⣣⣼⢿⣿⠿⠋⠀⠀⢻⡇⢀⠀⣀⠀⣿⡟
 */

#pragma once

#include <Wbemidl.h>
#include <atlsafe.h>
#include <comdef.h>

#include <iostream>
#include <iterator>
#include <memory>
#include <optional>
#include <stdexcept>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#pragma comment(lib, "wbemuuid.lib")

namespace wmi {

class Interface;

/**
 * custom exception type for wmi operations that provides clear error context
 * wraps standard runtime_error to maintain compatibility while adding wmi-specific context
 */
struct Exception final : std::runtime_error {
    explicit Exception(const std::string& message) : std::runtime_error(message) {}
};

/**
 * formats hresult error codes into human-readable strings for debugging
 * combines operation description with hex error code for quick diagnosis
 * \param operation - description of what failed
 * \param hr - the hresult error code from windows api
 * \returns formatted error string with operation context
 */
[[nodiscard]] inline std::string FormatHResultError(const std::string& operation, HRESULT hr) {
    return operation + " (HRESULT: 0x" + std::to_string(static_cast<unsigned>(hr)) + ")";
}

/**
 * converts com variant containing bstr to standard string safely
 * handles null bstr values and type checking to prevent crashes
 * \param variant - com variant that may contain bstr data
 * \returns optional string if conversion succeeds, nullopt otherwise
 */
[[nodiscard]] inline std::optional<std::string> ConvertBstrToString(const CComVariant& variant) {
    if (variant.vt == VT_BSTR && variant.bstrVal) {
        return std::string(_com_util::ConvertBSTRToString(variant.bstrVal));
    }
    return std::nullopt;
}

/**
 * manages com library initialization and cleanup using raii pattern
 * ensures proper com initialization on construction and cleanup on destruction
 * prevents resource leaks by automatically calling couninitialize when needed
 */
class COMInitializer {
   public:
    explicit COMInitializer(DWORD threading_model = COINIT_MULTITHREADED) : initialized_(false) {
        HRESULT hr = CoInitializeEx(nullptr, threading_model);
        if (SUCCEEDED(hr)) {
            initialized_ = true;
        } else if (hr != RPC_E_CHANGED_MODE) {
            // jgn dulu plis
            throw Exception(FormatHResultError("Failed to initialize COM library", hr));
        }

        hr = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT,
                                  RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);

        if (FAILED(hr) && hr != RPC_E_TOO_LATE) {
            // YYYYYYYYY
            if (initialized_) {
                CoUninitialize();
                initialized_ = false;
            }
            throw Exception(FormatHResultError("Failed to initialize COM security", hr));
        }
    }

    ~COMInitializer() noexcept {
        if (initialized_) {
            CoUninitialize();
        }
    }

    COMInitializer(const COMInitializer&) = delete;
    COMInitializer& operator=(const COMInitializer&) = delete;
    COMInitializer(COMInitializer&&) = delete;
    COMInitializer& operator=(COMInitializer&&) = delete;

    [[nodiscard]] bool IsInitialized() const noexcept { return initialized_; }

   private:
    bool initialized_;
};

/**
 * generic converter for com variants to c++ types with comprehensive error handling
 * handles type conversion failures gracefully and provides detailed error logging
 * \tparam T - target c++ type for conversion
 * \param variant - com variant to convert from
 * \returns optional containing converted value or nullopt on failure
 */
template <typename T>
[[nodiscard]] std::optional<T> ConvertVariant(const CComVariant& variant) {
    try {
        return static_cast<T>(variant_t(variant));
    } catch (const _com_error& e) {
        // error com
        std::cerr << "ConvertVariant failed: COM error 0x" << std::hex << e.Error() << " - "
                  << e.ErrorMessage() << std::endl;
        return std::nullopt;
    } catch (const std::exception& e) {
        // error std
        std::cerr << "ConvertVariant failed: " << e.what() << std::endl;
        return std::nullopt;
    } catch (...) {
        // ga jelas njir
        std::cerr << "ConvertVariant failed: Unknown error" << std::endl;
        return std::nullopt;
    }
}

/**
 * specialized converter for string variants using optimized bstr conversion
 * delegates to dedicated bstr converter for consistent string handling
 * \param variant - com variant expected to contain bstr data
 * \returns optional string if variant contains valid bstr data
 */
template <>
[[nodiscard]] inline std::optional<std::string> ConvertVariant<std::string>(
    const CComVariant& variant) {
    return ConvertBstrToString(variant);
}

/**
 * specialized converter for string arrays from com safe arrays
 * handles bstr array conversion to std::vector<std::string> with proper memory management
 * \param variant - com variant expected to contain safe array of bstr strings
 * \returns vector of strings if variant contains valid bstr array, nullopt otherwise
 */
template <>
[[nodiscard]] inline std::optional<std::vector<std::string>>
ConvertVariant<std::vector<std::string>>(const CComVariant& variant) {
    //array string
    if (variant.vt != (VT_ARRAY | VT_BSTR)) {
        // no array string
        return std::nullopt;
    }

    CComSafeArray<BSTR> safe_array;
    try {
        safe_array.Attach(variant.parray);
    } catch (...) {
        return std::nullopt;
    }

    std::vector<std::string> result;
    const ULONG count = safe_array.GetCount();
    result.reserve(count);

    for (ULONG i = 0; i < count; ++i) {
        if (auto bstr = safe_array.GetAt(i)) {
            result.emplace_back(_com_util::ConvertBSTRToString(bstr));
            // hell yeahhh
        }
    }

    safe_array.Detach();
    return result; // YEAHHHHHHHHHHHHHHHHHHHHH
}

/**
 * represents a single wmi object with property access capabilities
 * provides type-safe property retrieval from wmi class instances
 * manages underlying com object lifecycle automatically
 */
class Object {
    friend class QueryResult;

   protected:
    Object(std::shared_ptr<const Interface> iface, CComPtr<IWbemClassObject> object)
        : iface_(std::move(iface)), object_(std::move(object)) {}

   public:
    /**
     * retrieves a property value from the wmi object with type conversion
     * automatically handles com variant conversion to requested c++ type
     * \tparam T - target type for property value (defaults to variant_t)
     * \param name - property name as wide string view
     * \returns optional containing property value if available and convertible
     */
    template <typename T = variant_t>
    [[nodiscard]] std::optional<T> GetProperty(const std::wstring_view name) const {
        CComVariant variant;
        const auto result = object_->Get(name.data(), 0, &variant, nullptr, nullptr);

        if (FAILED(result)) {
            return std::nullopt;
        }

        if constexpr (std::is_same_v<T, variant_t>) {
            return variant;
        }

        return ConvertVariant<T>(variant);
    }

   private:
    std::shared_ptr<const Interface> iface_;
    CComPtr<IWbemClassObject> object_;
};

/**
 * represents the result of a wmi query with iterator-based access
 * provides lazy evaluation of query results through input iterators
 * automatically manages enumerator lifecycle and batch loading
 */
class QueryResult {
    friend class Interface;

   public:
    class Iterator {
       public:
        using iterator_category = std::input_iterator_tag;
        using value_type = Object;
        using difference_type = std::ptrdiff_t;
        using pointer = const Object*;
        using reference = const Object&;

        static constexpr ULONG BATCH_SIZE = 10;

        Iterator(std::shared_ptr<const Interface> iface, CComPtr<IEnumWbemClassObject> enumerator,
                 bool is_end = false)
            : iface_(std::move(iface)),
              enumerator_(std::move(enumerator)),
              is_end_(is_end),
              current_index_(0) {
            if (!is_end_ && enumerator_) {
                FetchNextBatch();
            }
        }

        reference operator*() const { return batch_[current_index_]; }

        pointer operator->() const { return &batch_[current_index_]; }

        Iterator& operator++() {
            if (is_end_) {
                return *this;
            }

            ++current_index_;
            if (current_index_ >= batch_.size()) {
                FetchNextBatch();
            }

            return *this;
        }

        Iterator operator++(int) {
            Iterator temp = *this;
            ++(*this);
            return temp;
        }

        bool operator==(const Iterator& other) const { return is_end_ == other.is_end_; }

        bool operator!=(const Iterator& other) const { return !(*this == other); }

       private:
        std::shared_ptr<const Interface> iface_;
        CComPtr<IEnumWbemClassObject> enumerator_;
        std::vector<Object> batch_;
        std::size_t current_index_;
        bool is_end_ = false;

        void FetchNextBatch() {
            batch_.clear();
            current_index_ = 0;

            if (!enumerator_) {
                is_end_ = true;
                return;
            }

            ULONG returned_count = 0;
            CComPtr<IWbemClassObject> objects[BATCH_SIZE];

            const auto result =
                enumerator_->Next(WBEM_INFINITE, BATCH_SIZE,
                                  reinterpret_cast<IWbemClassObject**>(objects), &returned_count);

            if (FAILED(result) || returned_count == 0) {
                is_end_ = true;
                return;
            }

            batch_.reserve(returned_count);
            for (ULONG i = 0; i < returned_count; ++i) {
                batch_.emplace_back(Object(iface_, objects[i]));
            }
        }
    };

   protected:
    QueryResult(std::shared_ptr<const Interface> iface,
                const CComPtr<IEnumWbemClassObject>& enumerator)
        : iface_(std::move(iface)), enumerator_(enumerator) {}

   public:
    /**
     * returns iterator to beginning of query results
     * resets enumerator to ensure consistent iteration from start
     * \returns iterator pointing to first result object
     */
    [[nodiscard]] Iterator begin() const {
        if (enumerator_) {
            enumerator_->Reset();
        }
        return Iterator(iface_, enumerator_);
    }

    /**
     * returns iterator to end of query results
     * creates sentinel iterator for range-based for loops
     * \returns end iterator marker
     */
    [[nodiscard]] Iterator end() const { return Iterator(iface_, nullptr, true); }

   private:
    std::shared_ptr<const Interface> iface_;
    CComPtr<IEnumWbemClassObject> enumerator_;
};

/**
 * main interface for wmi operations providing namespace connection and query execution
 * manages underlying com connections and provides thread-safe query capabilities
 * uses shared_ptr for automatic lifetime management across threads
 */
class Interface : public std::enable_shared_from_this<const Interface> {
   public:
    struct PassKey {
        explicit PassKey() = default;
    };

    /**
     * factory method to create interface instances with proper initialization
     * \param path - wmi namespace path (defaults to "cimv2")
     * \returns shared_ptr to initialized interface ready for queries
     */
    static std::shared_ptr<Interface> Create(std::string_view path = "cimv2");

    explicit Interface(PassKey, const std::string_view path) {
        auto result = CoCreateInstance(CLSID_WbemLocator, nullptr, CLSCTX_INPROC_SERVER,
                                       IID_IWbemLocator, reinterpret_cast<LPVOID*>(&locator_));
        if (FAILED(result)) {
            // error locator
            throw Exception("Failed to create WbemLocator object. " +
                            FormatHResultError("Check if WMI service is available", result));
        }

        // wmi namespace
        result = locator_->ConnectServer(bstr_t(R"(\\.\root\)") + bstr_t(path.data()), nullptr,
                                         nullptr, nullptr, 0, nullptr, nullptr, &services_);
        if (FAILED(result)) {
            throw Exception(
                "Could not connect to WMI namespace '" + std::string(path) + "'. " +
                FormatHResultError("Verify namespace exists and access permissions", result));
        }

        // apcb
        result = CoSetProxyBlanket(services_, RPC_C_AUTHN_DEFAULT, RPC_C_AUTHZ_NONE,
                                   COLE_DEFAULT_PRINCIPAL, RPC_C_AUTHN_LEVEL_DEFAULT,
                                   RPC_C_IMP_LEVEL_IMPERSONATE, nullptr, EOAC_NONE);
        if (FAILED(result)) {
            // error security
            throw Exception("Could not set proxy blanket for WMI connection. " +
                            FormatHResultError("Authentication may have failed", result));
        }
    }

    ~Interface() noexcept = default;

    Interface(const Interface& other) = delete;
    Interface& operator=(const Interface& other) = delete;

    Interface(Interface&& other) noexcept = default;
    Interface& operator=(Interface&& other) noexcept = default;

    /**
     * executes wql query against connected wmi namespace
     * provides lazy-evaluated results through queryresult iterator interface
     * \param query - wql query string as wide character view
     * \returns queryresult object for iterating over matching wmi objects
     * \throws Exception if query execution fails with detailed error context
     */
    [[nodiscard]] QueryResult ExecuteQuery(const std::wstring_view query) const {
        CComPtr<IEnumWbemClassObject> enumerator;
        const auto query_bstr = bstr_t(query.data());
        const auto result = services_->ExecQuery(
            bstr_t("WQL"), query_bstr, WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            nullptr, &enumerator);

        if (FAILED(result)) {
            const std::string query_str = _com_util::ConvertBSTRToString(query_bstr);
            throw Exception(
                "WQL query execution failed for query: '" + query_str + "'. " +
                FormatHResultError("Check query syntax and target class availability", result));
        }

        return {shared_from_this(), enumerator};
    }

   private:
    CComPtr<IWbemLocator> locator_;
    CComPtr<IWbemServices> services_;
};

/**
 * factory implementation that creates interface instances through make_shared
 * ensures proper memory management and exception safety during construction
 * \param path - wmi namespace to connect to
 * \returns shared_ptr to fully initialized interface instance
 */
inline std::shared_ptr<Interface> Interface::Create(const std::string_view path) {
    return std::make_shared<Interface>(PassKey{}, path);
}

}  // namespace wmi