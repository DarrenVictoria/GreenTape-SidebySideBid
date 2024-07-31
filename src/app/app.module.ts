import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';


import { AppComponent } from './app.component';
import { SidebySideBidComponent } from './sidebyside-bid/sidebyside-bid.component';


const routes: Routes = [
  { path: 'side', component: SidebySideBidComponent },
  { path: '', redirectTo: '/side', pathMatch: 'full' },

];

@NgModule({
  declarations: [
    AppComponent,
    SidebySideBidComponent
  ],
  imports: [
    BrowserModule,
    RouterModule.forRoot(routes)
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
