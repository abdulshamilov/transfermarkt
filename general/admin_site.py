from django.contrib.admin import AdminSite
from django.contrib.auth.models import User, Group
from django.contrib.auth.admin import UserAdmin, GroupAdmin


class TransfermarktAdminSite(AdminSite):
    site_header = 'Transfermarkt'
    site_title = 'Transfermarkt Admin'
    index_title = 'Dashboard'

    def get_app_list(self, request, app_label=None):
        app_list = super().get_app_list(request, app_label)
        app_list.sort(key=lambda app: app['app_label'] != 'general')
        return app_list

    def index(self, request, extra_context=None):
        from .models import Player, Club, League, National, Tournament

        extra_context = extra_context or {}
        extra_context['stats'] = [
            {'label': 'Players', 'value': Player.objects.count(), 'icon': 'players'},
            {'label': 'Clubs', 'value': Club.objects.count(), 'icon': 'clubs'},
            {'label': 'Leagues', 'value': League.objects.count(), 'icon': 'leagues'},
            {'label': 'Tournaments', 'value': Tournament.objects.count(), 'icon': 'tournaments'},
        ]
        extra_context['nation_count'] = National.objects.count()
        return super().index(request, extra_context)


site = TransfermarktAdminSite(name='admin')
site.register(User, UserAdmin)
site.register(Group, GroupAdmin)
