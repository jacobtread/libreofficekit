#include "wrapper.hpp"

LibreOfficeKit *
lok_init_wrapper(const char *install_path)
{
  return lok_init(install_path);
}