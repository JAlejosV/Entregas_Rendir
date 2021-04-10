using System;

namespace EntregasRendir
{
    class Program
    {
        static void Main()
        {
            EntregasRendir entregasRendir = new EntregasRendir();
            try
            {
                Logger.WriteLine("Inicio de Migracion de las ER");
                entregasRendir.EntregasRendirMigracionHelmAOfisis();
                Logger.WriteLine("Fin de Migracion de las ER");
            }
            catch (Exception ex)
            {
                Logger.WriteLine($"{ex.Message}\n{ex.InnerException?.Message}");
                throw;
            }

        }
    }
}
