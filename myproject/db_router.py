class MultiDBRouter:
    def db_for_write(self, model, **hints):
        # Por defecto escribe en SQLite
        return 'default'

    def db_for_read(self, model, **hints):
        return 'default'

    def allow_relation(self, obj1, obj2, **hints):
        return True

    def allow_migrate(self, db, app_label, model_name=None, **hints):
        # Migrar todas las apps a ambas bases
        return True
