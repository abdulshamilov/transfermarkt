from django.contrib import admin
from .models import National, Player, League, Club, Tournament, TournamentParticipant
from .admin_site import site


@admin.register(Player, site=site)
class PlayerModelAdmin(admin.ModelAdmin):
    list_display = (
        'name',
        'position',
        'club',
        'get_age',
        'games',
        'goals',
        'assists',
        'price',
    )
    search_fields = ('name', 'club__title', 'league__title')
    list_filter = ('position', 'club', 'league')
    ordering = ('-price',)

    def get_age(self, obj):
        return obj.get_age()
    get_age.short_description = 'Age'


@admin.register(League, site=site)
class LeagueModelAdmin(admin.ModelAdmin):
    list_display = (
        'title',
        'earth',
    )
    search_fields = ('title',)
    list_filter = ('earth',)


@admin.register(Club, site=site)
class ClubModelAdmin(admin.ModelAdmin):
    list_display = (
        'title',
        'league',
        'price',
    )
    search_fields = ('title',)
    list_filter = ('league',)
    ordering = ('-price',)


@admin.register(National, site=site)
class NationalModelAdmin(admin.ModelAdmin):
    list_display = (
        'title',
        'continent',
    )
    search_fields = ('title',)
    list_filter = ('continent',)


@admin.register(Tournament, site=site)
class TournamentModelAdmin(admin.ModelAdmin):
    list_display = (
        'name',
        'season',
        'is_national',
    )
    search_fields = ('name',)
    list_filter = ('season', 'is_national')


@admin.register(TournamentParticipant, site=site)
class TournamentParticipantAdmin(admin.ModelAdmin):
    list_display = (
        'tournament',
        'club',
        'points',
        'position',
    )
    search_fields = ('tournament__name', 'club__title')
    list_filter = ('tournament',)
